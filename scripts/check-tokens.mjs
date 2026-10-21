#!/usr/bin/env node

/**
 * scripts/check-tokens.mjs
 * 视觉规则闸门检查脚本 (P6 规范)
 * 扫描 web/src 下所有 .svelte 的 <style> 块和 .css 文件
 *
 * 检查项：
 * 1. 裸颜色：#hex / rgb() / rgba()（白名单仅限 web/src/styles/office-fluent.css 的 :root token 定义区）
 * 2. 禁用字号：font-size 里的 9px / 10px / 10.5px / 11.5px（11px 允许，单独统计）
 * 3. 圆角：border-radius 里除 4px / 6px / 999px / var(--office-radius*) 之外的值
 * 4. emoji：模板区的 emoji 码点（U+1F300–U+1FAFF、U+2600–U+27BF）
 * 5. 原生弹窗：alert( / confirm(
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 根目录与目标扫描目录
const projectRoot = path.resolve(__dirname, '..');
const targetDir = path.resolve(projectRoot, 'web/src');

function getAllFiles(dir, extensions) {
  let results = [];
  if (!fs.existsSync(dir)) return results;
  const list = fs.readdirSync(dir);
  for (const file of list) {
    const fullPath = path.join(dir, file);
    const stat = fs.statSync(fullPath);
    if (stat.isDirectory()) {
      results = results.concat(getAllFiles(fullPath, extensions));
    } else {
      if (extensions.some(ext => file.endsWith(ext))) {
        results.push(fullPath);
      }
    }
  }
  return results;
}

const targetFiles = getAllFiles(targetDir, ['.svelte', '.css']);

let rawColorCount = 0;
let smallFontCount = 0;
let font11Count = 0;
let borderRadiusCount = 0;
let emojiCount = 0;
let nativePopupCount = 0;

const findings = [];

// 规则正则定义
const emojiRegex = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u;
const nativePopupRegex = /\b(alert|confirm)\s*\(/;
const rawColorHexRegex = /#([0-9a-fA-F]{3,8})\b/g;
const rawColorRgbRegex = /rgba?\([^)]+\)/g;
const fontSizeRegex = /font-size\s*:\s*([^;]+)/i;
const borderRadiusRegex = /border-radius\s*:\s*([^;]+)/i;

// 允许的圆角白名单
const allowedRadiusPatterns = [
  /^0$/,
  /^4px$/,
  /^6px$/,
  /^999px$/,
  /^9999px$/,
  /^50%$/,
  /^inherit$/,
  /^initial$/,
  /^var\(--office-radius[^)]*\)$/
];

for (const filePath of targetFiles) {
  const relativePath = path.relative(projectRoot, filePath).replace(/\\/g, '/');
  const isCss = filePath.endsWith('.css');
  const isOfficeFluentCss = relativePath.endsWith('web/src/styles/office-fluent.css');
  const content = fs.readFileSync(filePath, 'utf8');
  const lines = content.split(/\r?\n/);

  let inStyleBlock = isCss;
  let inTokenDefSection = false;

  lines.forEach((line, index) => {
    const lineNum = index + 1;
    const trimmed = line.trim();

    // Svelte style 块范围跟踪
    if (!isCss) {
      if (line.includes('<style')) {
        inStyleBlock = true;
      }
      if (line.includes('</style>')) {
        inStyleBlock = false;
        return;
      }
    }

    // office-fluent.css :root token 定义区白名单跟踪
    if (isOfficeFluentCss) {
      if (trimmed.startsWith(':root')) {
        inTokenDefSection = true;
      }
      if (inTokenDefSection && trimmed === '}') {
        inTokenDefSection = false;
        return;
      }
    }

    // 1. 裸颜色检查 (仅在样式区生效，排除 token 定义文件 office-fluent.css 与注释)
    if (inStyleBlock && !inTokenDefSection && !isOfficeFluentCss) {
      // 简单忽略纯注释行
      const codeOnly = trimmed.startsWith('/*') && trimmed.endsWith('*/') ? '' : trimmed;
      if (codeOnly) {
        const hexMatches = codeOnly.match(rawColorHexRegex) || [];
        const rgbMatches = codeOnly.match(rawColorRgbRegex) || [];
        const allColors = [...hexMatches, ...rgbMatches];
        for (const col of allColors) {
          rawColorCount++;
          findings.push(`${relativePath}:${lineNum} 裸色 ${col}`);
        }
      }
    }

    // 2. 禁用字号检查 (在样式区生效)
    if (inStyleBlock) {
      const fsMatch = trimmed.match(fontSizeRegex);
      if (fsMatch) {
        const val = fsMatch[1].trim();
        if (/\b(9px|10px|10\.5px|11\.5px)\b/i.test(val)) {
          smallFontCount++;
          findings.push(`${relativePath}:${lineNum} 小字 ${val}`);
        }
        if (/\b11px\b/i.test(val)) {
          font11Count++;
        }
      }
    }

    // 3. 圆角检查 (在样式区生效)
    if (inStyleBlock) {
      const brMatch = trimmed.match(borderRadiusRegex);
      if (brMatch) {
        const val = brMatch[1].trim();
        const isAllowed = allowedRadiusPatterns.some(pat => pat.test(val));
        if (!isAllowed) {
          borderRadiusCount++;
          findings.push(`${relativePath}:${lineNum} 圆角 ${val}`);
        }
      }
    }

    // 4. Emoji 检查 (全文件或模板区)
    const emojiMatch = line.match(emojiRegex);
    if (emojiMatch) {
      emojiCount++;
      findings.push(`${relativePath}:${lineNum} emoji ${emojiMatch[0]}`);
    }

    // 5. 原生弹窗检查 (全文件)
    const popupMatch = line.match(nativePopupRegex);
    if (popupMatch) {
      nativePopupCount++;
      findings.push(`${relativePath}:${lineNum} 原生弹窗 ${popupMatch[0]}`);
    }
  });
}

// 输出详细检查项 (每条一行: 文件:行 规则 内容)
for (const item of findings) {
  console.log(item);
}

// 结尾汇总输出
console.log(
  `\n裸色 ${rawColorCount} / 小字 ${smallFontCount} / 圆角 ${borderRadiusCount} / emoji ${emojiCount} / 原生弹窗 ${nativePopupCount}` +
  (font11Count > 0 ? ` (合规 11px 字号: ${font11Count})` : '')
);

// 模式控制：根据 P6 规范，先只做 warn（退出码 0），待存量清理确认后再改为 exit 1
const isStrictMode = process.argv.includes('--strict') || process.env.CHECK_STRICT === '1';
const totalViolations = rawColorCount + smallFontCount + borderRadiusCount + emojiCount + nativePopupCount;

if (totalViolations > 0 && isStrictMode) {
  console.error(`[FAIL] 视觉规则检查未通过，共发现 ${totalViolations} 处违规。`);
  process.exit(1);
} else {
  if (totalViolations > 0) {
    console.log(`[WARN] 视觉规则检查完成 (当前为 warn 模式，退出码 0)。`);
  } else {
    console.log(`[PASS] 视觉规则检查 100% 完美通过！`);
  }
  process.exit(0);
}
