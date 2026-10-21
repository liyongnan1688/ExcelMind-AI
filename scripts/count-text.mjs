#!/usr/bin/env node

/**
 * scripts/count-text.mjs
 * 界面文案与汉字量复现统计脚本
 * 
 * 验收统一口径：
 * 扫描 web/src 下所有 .svelte 组件模板区（排除 <script> 与 <style>）：
 * 1. 静态纯文本节点 (Text Nodes，剥离 HTML 标签与花括号插值表达式)
 * 2. 交互提示属性 (placeholder, title, label)
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const webSrcDir = path.resolve(projectRoot, 'web/src');

function getAllSvelteFiles(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  for (const file of list) {
    const fullPath = path.join(dir, file);
    const stat = fs.statSync(fullPath);
    if (stat.isDirectory()) {
      results = results.concat(getAllSvelteFiles(fullPath));
    } else if (file.endsWith('.svelte')) {
      results.push(fullPath);
    }
  }
  return results.sort();
}

function countChineseChars(text) {
  if (!text) return 0;
  const matches = text.match(/[\u4e00-\u9fa5]/g);
  return matches ? matches.length : 0;
}

const svelteFiles = getAllSvelteFiles(webSrcDir);

console.log('======================================================================');
console.log('              ExcelMind AI 界面模板区汉字统计报告                     ');
console.log('======================================================================\n');

let totalStaticText = 0;
let totalAttributes = 0;

const rows = [];

for (const filePath of svelteFiles) {
  const relPath = path.relative(projectRoot, filePath).replace(/\\/g, '/');
  const content = fs.readFileSync(filePath, 'utf8');

  // 剔除 <script> 与 <style> 块
  const templateOnly = content
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '');

  // 1. 交互属性提取 (严格匹配 placeholder, title, label，排除 aria-label)
  let attrText = '';
  const attrRegex = /(?<![\w-])(placeholder|title|label)\s*=\s*(?:"([^"]*)"|'([^']*)')/gi;
  let match;
  while ((match = attrRegex.exec(templateOnly)) !== null) {
    attrText += (match[2] || match[3] || '') + ' ';
  }
  const attrCount = countChineseChars(attrText);

  // 2. 静态纯文本节点 (剥离 HTML 标签与花括号插值)
  let staticText = templateOnly.replace(/<[^>]+>/g, ' ');
  let prev;
  do {
    prev = staticText;
    staticText = staticText.replace(/\{[^{}]*\}/g, ' ');
  } while (staticText !== prev);
  const staticCount = countChineseChars(staticText);

  rows.push({
    file: relPath,
    staticCount,
    attrCount,
    combined: staticCount + attrCount
  });

  totalStaticText += staticCount;
  totalAttributes += attrCount;
}

console.log('文件路径                                     | 静态文本 | 属性文案 | 合计汉字');
console.log('--------------------------------------------|----------|----------|----------');
for (const r of rows) {
  const fileCol = r.file.padEnd(43);
  const staticCol = String(r.staticCount).padStart(8);
  const attrCol = String(r.attrCount).padStart(8);
  const combinedCol = String(r.combined).padStart(8);
  console.log(`${fileCol} | ${staticCol} | ${attrCol} | ${combinedCol}`);
}
console.log('--------------------------------------------|----------|----------|----------');
const grandTotal = totalStaticText + totalAttributes;
console.log(
  `合计汇总 (${rows.length} 个组件)                       | ${String(totalStaticText).padStart(8)} | ${String(totalAttributes).padStart(8)} | ${String(grandTotal).padStart(8)}\n`
);

console.log('【口径说明与降幅核算】:');
console.log(`- 本脚本实测总计（静态纯文本 + 交互属性）: ${grandTotal} 汉字`);
console.log(`  (纯文本节点: ${totalStaticText} 字，属性节点: ${totalAttributes} 字)`);
const narrowDiff = grandTotal - 6287;
console.log(`- 与静态收窄口径（6287）相差 ${narrowDiff} 字，差异主要来自插值节点内字面量及特定属性的收敛范围。`);
const baseline = 8586;
const reduction = (((baseline - grandTotal) / baseline) * 100).toFixed(1);
console.log(`- 基线（commit c1c18d1）为 ${baseline} 字，当前为 ${grandTotal} 字，降幅为 -${reduction}%（尚未达到 4500 字目标）。\n`);
