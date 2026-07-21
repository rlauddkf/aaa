// 오프라인 자체 검증: 합성 일봉 데이터로 지표/스코어/리포트 파이프라인을 점검
const { computeMetrics } = require('./indicators');
const { scoreSymbol } = require('./score');
const { formatReport } = require('./recommend');

// 합성 일봉 생성기: 추세 강도(trend)와 변동성(vol)을 조절
function genBars(n, { start = 100, trend = 0, vol = 0.02, baseVol = 1e7, volSpike = 1 } = {}) {
  const bars = [];
  let price = start;
  let seed = 42;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  for (let i = 0; i < n; i++) {
    const drift = trend;
    const shock = (rnd() - 0.5) * 2 * vol;
    const open = price;
    const close = Math.max(1, price * (1 + drift + shock));
    const high = Math.max(open, close) * (1 + rnd() * vol);
    const low = Math.min(open, close) * (1 - rnd() * vol);
    const spike = i >= n - 1 ? volSpike : 1;
    const volume = Math.round(baseVol * (0.7 + rnd() * 0.6) * spike);
    bars.push({ date: `2026-01-${String((i % 28) + 1).padStart(2, '0')}`, open, high, low, close, volume });
    price = close;
  }
  return bars;
}

const cases = [
  ['STRONG_UP', genBars(60, { trend: 0.006, vol: 0.03, baseVol: 3e7, volSpike: 2.2 })],
  ['CHOPPY',    genBars(60, { trend: 0.0,   vol: 0.008, baseVol: 2e7, volSpike: 1 })],
  ['DOWN_VOL',  genBars(60, { trend: -0.005, vol: 0.04, baseVol: 5e7, volSpike: 1.8 })],
  ['LOW_LIQ',   genBars(60, { trend: 0.004, vol: 0.03, baseVol: 2e5, volSpike: 1 })], // 유동성 필터 대상
  ['CALM_BIG',  genBars(60, { trend: 0.001, vol: 0.006, baseVol: 8e7, volSpike: 1.1 })],
];

console.log('=== 지표/스코어 검증 ===');
const scored = [];
for (const [name, bars] of cases) {
  const m = computeMetrics(bars);
  const s = scoreSymbol(name, m);
  if (!s) { console.log(`${name}: metrics 계산 실패`); continue; }
  if (s.filtered) { console.log(`${name}: [필터됨] ${s.reason} (거래대금 $${(m.dollarVol/1e6).toFixed(1)}M)`); continue; }
  s.isETF = false;
  scored.push(s);
  console.log(`${name.padEnd(10)} score=${String(s.score).padStart(5)} bias=${s.bias.split('(')[0]} ` +
    `ATR%=${m.atrPct.toFixed(2)} RVOL=${m.rvol.toFixed(2)} RSI=${m.rsi14.toFixed(0)} ` +
    `mom5=${m.mom5.toFixed(1)} $Vol=${(m.dollarVol/1e6).toFixed(0)}M`);
}

scored.sort((a, b) => b.score - a.score);
const result = {
  generatedAt: new Date().toISOString(),
  universeSize: cases.length, fetched: cases.length, scoredCount: scored.length,
  errors: [], ranking: scored.slice(0, 5), pick: scored[0] || null,
};

console.log('\n=== 리포트 출력 검증 ===');
console.log(formatReport(result));

// 간단한 단언
const names = scored.map((s) => s.symbol);
if (names.includes('LOW_LIQ')) { console.error('\n❌ FAIL: 저유동성 종목이 필터되지 않음'); process.exit(1); }
if (!result.pick) { console.error('\n❌ FAIL: 추천 종목 없음'); process.exit(1); }
if (result.pick.score < 0 || result.pick.score > 100) { console.error('\n❌ FAIL: 점수 범위 오류'); process.exit(1); }
console.log('\n✅ PASS: 필터/스코어/리포트 정상 동작');
