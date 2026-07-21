#!/usr/bin/env node
// 매일 저녁 5시 30분(기본 KST)에 자동으로 추천을 생성하는 스케줄러
//   사용법:  node stock/scheduler.js       (계속 실행되며 매일 17:30에 동작)
//           npm run schedule
//
// 환경변수:
//   REC_CRON   크론 표현식 (기본 "30 17 * * *" = 매일 17:30)
//   REC_TZ     타임존       (기본 "Asia/Seoul")
//   RUN_NOW    "1" 이면 시작 시 즉시 1회 실행 후 스케줄 대기

const cron = require('node-cron');
const { runRecommendation, formatReport, saveReport } = require('./recommend');

const CRON = process.env.REC_CRON || '30 17 * * *'; // 매일 17:30
const TZ = process.env.REC_TZ || 'Asia/Seoul';

async function job(reason = 'scheduled') {
  const started = new Date().toLocaleString('ko-KR', { timeZone: TZ });
  console.log(`\n[${started}] ▶ 추천 생성 시작 (${reason})`);
  try {
    const result = await runRecommendation();
    const report = formatReport(result);
    console.log('\n' + report + '\n');
    const { txtPath } = saveReport(result, report);
    console.log(`💾 저장됨: ${txtPath}`);
  } catch (e) {
    console.error('❌ 추천 생성 실패:', e.message);
  }
}

if (!cron.validate(CRON)) {
  console.error(`❌ 잘못된 크론 표현식: ${CRON}`);
  process.exit(1);
}

cron.schedule(CRON, () => job('cron'), { timezone: TZ });

console.log('════════════════════════════════════════════');
console.log('  📈 미국장 단타 추천 스케줄러 실행 중');
console.log(`  스케줄 : ${CRON}  (${TZ})`);
console.log('  기본값 : 매일 저녁 5시 30분에 자동 분석');
console.log('  중지   : Ctrl + C');
console.log('════════════════════════════════════════════');

if (process.env.RUN_NOW === '1') {
  job('startup(RUN_NOW)');
}

// 프로세스 유지
process.stdin.resume();
