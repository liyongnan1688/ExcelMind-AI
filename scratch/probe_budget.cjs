const fs = require('fs');
const https = require('https');

const logPath = process.env.LOCALAPPDATA + '\\LeeExcel\\WebView2Profile\\EBWebView\\Default\\Local Storage\\leveldb\\000003.log';
const logContent = fs.readFileSync(logPath, 'utf8');
const matches = logContent.match(/\{[^{}]*provider[^{}]*\}/g);
const config = JSON.parse(matches[matches.length - 1]);

async function testParam(payload) {
  return new Promise((resolve) => {
    const url = new URL(config.baseUrl.replace(/\/+$/, '') + '/chat/completions');
    const req = https.request(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + config.apiKey
      },
      timeout: 30000
    }, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        resolve({ status: res.statusCode, body });
      });
    });
    req.on('error', (e) => resolve({ error: e.message }));
    req.write(JSON.stringify(payload));
    req.end();
  });
}

(async () => {
  // Test max_tokens: 32768
  const r1 = await testParam({
    model: config.model,
    messages: [{ role: 'user', content: 'hi' }],
    max_tokens: 32768
  });
  console.log('max_tokens 32768 status:', r1.status);
  try {
    const j1 = JSON.parse(r1.body);
    if (j1.error) console.log('32768 error:', j1.error.message);
    else console.log('32768 success!');
  } catch(e) {}

  // Test thinking budget_tokens
  const r2 = await testParam({
    model: config.model,
    messages: [{ role: 'user', content: 'hi' }],
    max_tokens: 8192,
    thinking: { type: 'enabled', budget_tokens: 1024 }
  });
  console.log('thinking budget 1024 status:', r2.status);
  try {
    const j2 = JSON.parse(r2.body);
    if (j2.error) console.log('budget error:', j2.error.message);
    else console.log('budget success! usage:', JSON.stringify(j2.usage));
  } catch(e) {}
})();
