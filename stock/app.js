#!/usr/bin/env node
// exe 진입점 (pkg로 패키징되는 메인 파일)
//
// 실행 모드 (첫 번째 인자):
//   (없음) | once   : 지금 1회 분석 → 결과 출력 → 저장 → "키 입력 대기"(더블클릭용)
//   once-quiet      : 1회 분석 후 즉시 종료 (Windows 작업 스케줄러 자동 실행용)
//   schedule        : 매일 저녁 5시 30분(KST) 자동 실행, 창을 켜둔 채 대기
//
// 환경변수 REC_CRON / REC_TZ / RUN_NOW 는 scheduler 모드에서 그대로 적용됩니다.

const { runRecommendation, formatReport, saveReport } = require('./recommend');

function waitKey(msg = '\n엔터 키를 누르면 창이 닫힙니다...') {
  return new Promise((resolve) => {
    process.stdout.write(msg);
    process.stdin.resume();
    process.stdin.setEncoding('utf8');
    process.stdin.once('data', () => resolve());
  });
}

async function runOnce({ pause }) {
  console.log('⏳ 미국장 단타 후보 분석 중... (수십 초 소요될 수 있습니다)\n');
  try {
    const result = await runRecommendation();
    const report = formatReport(result);
    console.log(report + '\n');
    const { txtPath } = saveReport(result, report);
    console.log(`💾 저장됨: ${txtPath}`);
    if (pause) await waitKey();
    process.exit(result.pick ? 0 : 1);
  } catch (e) {
    console.error('❌ 실행 실패:', e.message);
    if (pause) await waitKey();
    process.exit(1);
  }
}

async function runSchedule() {
  const cron = require('node-cron');
  const CRON = process.env.REC_CRON || '30 17 * * *';
  const TZ = process.env.REC_TZ || 'Asia/Seoul';

  async function job(reason) {
    const t = new Date().toLocaleString('ko-KR', { timeZone: TZ });
    console.log(`\n[${t}] ▶ 추천 생성 시작 (${reason})`);
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
  console.log('  ※ 이 창을 켜두어야 자동 실행됩니다. 중지: Ctrl + C');
  console.log('════════════════════════════════════════════');

  if (process.env.RUN_NOW === '1') job('startup(RUN_NOW)');
  process.stdin.resume();
}

const mode = (process.argv[2] || 'once').toLowerCase();
if (mode === 'schedule') runSchedule();
else if (mode === 'once-quiet') runOnce({ pause: false });
else runOnce({ pause: true });
