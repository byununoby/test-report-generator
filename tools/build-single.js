// 모든 스크립트를 index.html 안에 넣어 "파일 1개"짜리 도구를 만든다.
// 문서관리시스템(ECM)이나 메일처럼 HTML 파일만 따로 열리는 환경에서도 동작하게 하기 위함.
// 사용: node tools/build-single.js <사이트 폴더> <출력 파일>
const fs = require('fs');
const path = require('path');

const dir = process.argv[2] || '.';
const out = process.argv[3] || path.join(dir, 'test-report-generator.html');
let html = fs.readFileSync(path.join(dir, 'index.html'), 'utf8');

// 인라인 스크립트를 허용하되, 외부 통신 차단(connect-src 'none')은 그대로 유지
html = html.replace("script-src 'self';", "script-src 'self' 'unsafe-inline';");

html = html.replace(/<script src="([^"]+)"><\/script>/g, (m, src) => {
  const file = path.join(dir, src);
  if (!fs.existsSync(file)) throw new Error('스크립트 파일이 없습니다: ' + src);
  const code = fs.readFileSync(file, 'utf8').replace(/<\/script/gi, '<\\/script');
  return '<script>/* ' + src + ' */\n' + code + '\n</script>';
});

if (/<script src=/.test(html)) throw new Error('인라인되지 않은 스크립트가 남아 있습니다.');
fs.writeFileSync(out, html);
console.log('만들었습니다:', out, Math.round(fs.statSync(out).size / 1024) + 'KB');
