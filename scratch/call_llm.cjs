const fs = require('fs');
const https = require('https');

const systemPromptFile = process.argv[2];
const userPrompt = process.argv[3];
const outputFile = process.argv[4];

if (!systemPromptFile || !userPrompt) {
  console.error("Usage: node call_llm.cjs <systemPromptFile> <userPrompt> [outputFile]");
  process.exit(1);
}

const systemPrompt = fs.readFileSync(systemPromptFile, 'utf8');

// 读取 LocalStorage 配置
const logPath = process.env.LOCALAPPDATA + '\\LeeExcel\\WebView2Profile\\EBWebView\\Default\\Local Storage\\leveldb\\000003.log';
const logContent = fs.readFileSync(logPath, 'utf8');
const matches = logContent.match(/\{[^{}]*provider[^{}]*\}/g);
if (!matches || matches.length === 0) {
  console.error("No config found");
  process.exit(1);
}

const config = JSON.parse(matches[matches.length - 1]);
const apiKey = config.apiKey;
const model = config.model || 'deepseek-flash';

const payload = JSON.stringify({
  model: model,
  messages: [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userPrompt }
  ],
  max_tokens: 8192
});

const req = https.request('https://api.deepseek.com/v1/chat/completions', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + apiKey
  },
  timeout: 180000
}, (res) => {
  let body = '';
  res.on('data', chunk => body += chunk);
  res.on('end', () => {
    try {
      const json = JSON.parse(body);
      const content = json.choices?.[0]?.message?.content || '';
      if (outputFile) {
        fs.writeFileSync(outputFile, content, 'utf8');
      } else {
        process.stdout.write(content);
      }
    } catch (err) {
      console.error("Error parsing response:", body);
      process.exit(1);
    }
  });
});

req.on('error', (e) => {
  console.error("HTTP request error:", e.message);
  process.exit(1);
});

req.write(payload);
req.end();
