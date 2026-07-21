#!/usr/bin/env node
// 즉시 1회 실행: 지금 바로 분석해서 추천 종목을 출력하고 logs/ 에 저장
//   사용법:  node stock/cli.js
//           npm run recommend

const { runRecommendation, formatReport, saveReport } = require('./recommend');

(async () => {
  console.log('⏳ 미국장 단타 후보 분석 중... (수십 초 소요될 수 있습니다)');
  try {
    const result = await runRecommendation();
    const report = formatReport(result);
    console.log('\n' + report + '\n');
    const { txtPath } = saveReport(result, report);
    console.log(`💾 저장됨: ${txtPath}`);
    process.exit(result.pick ? 0 : 1);
  } catch (e) {
    console.error('❌ 실행 실패:', e.message);
    process.exit(1);
  }
})();
