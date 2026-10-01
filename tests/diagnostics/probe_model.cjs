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
  // Test 1: Simple greeting with max_tokens: 100
  const r1 = await testParam({
    model: config.model,
    messages: [{ role: 'user', content: 'hi' }],
    max_tokens: 100
  });
  console.log('Test 1 (max_tokens 100) status:', r1.status);
  try {
    const j1 = JSON.parse(r1.body);
    console.log('Test 1 model returned:', j1.model);
    console.log('Test 1 has reasoning_content?', !!j1.choices?.[0]?.message?.reasoning_content);
    console.log('Test 1 finish_reason:', j1.choices?.[0]?.finish_reason);
    console.log('Test 1 usage:', JSON.stringify(j1.usage));
  } catch(e) { console.log('Test 1 raw:', r1.body?.substring(0, 200)); }

  // Test 2: Try max_tokens: 16384
  const r2 = await testParam({
    model: config.model,
    messages: [{ role: 'user', content: 'hi' }],
    max_tokens: 16384
  });
  console.log('Test 2 (max_tokens 16384) status:', r2.status);
  try {
    const j2 = JSON.parse(r2.body);
    if (j2.error) console.log('Test 2 error:', j2.error);
    else console.log('Test 2 success, max_tokens 16384 accepted! usage:', JSON.stringify(j2.usage));
  } catch(e) { console.log('Test 2 raw:', r2.body?.substring(0, 200)); }

  // Test 3: Try max_tokens: 8192 with thinking: { type: "disabled" } or reasoning_effort: "low"
  const r3 = await testParam({
    model: config.model,
    messages: [{ role: 'user', content: 'hi' }],
    max_tokens: 200,
    thinking: { type: 'disabled' }
  });
  console.log('Test 3 (thinking: disabled) status:', r3.status);
  try {
    const j3 = JSON.parse(r3.body);
    if (j3.error) console.log('Test 3 error message:', j3.error.message);
    else {
      console.log('Test 3 success! has reasoning_content?', !!j3.choices?.[0]?.message?.reasoning_content);
      console.log('Test 3 usage:', JSON.stringify(j3.usage));
    }
  } catch(e) {}
})();
