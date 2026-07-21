// 추천 오케스트레이터
// 유니버스 전체를 수집 → 지표 계산 → 스코어링 → 1위 선정 → 리포트 생성/저장

const fs = require('fs');
const path = require('path');
const { UNIVERSE, ETFS } = require('./universe');
const { fetchAll } = require('./dataSource');
const { computeMetrics } = require('./indicators');
const { scoreSymbol } = require('./score');

const ETF_SET = new Set(ETFS);

// 로그 저장 기준 폴더.
// - exe(pkg)로 실행 시: 스냅샷 내부는 읽기전용이므로 exe 파일이 있는 폴더 사용
// - 일반 node 실행 시: 프로젝트 루트
function baseDir() {
  if (process.pkg) return path.dirname(process.execPath);
  return path.join(__dirname, '..');
}

function fmt(n, d = 2) {
  return n == null || Number.isNaN(n) ? 'N/A' : Number(n).toFixed(d);
}

function humanVol(v) {
  if (v == null) return 'N/A';
  if (v >= 1e9) return (v / 1e9).toFixed(2) + 'B';
  if (v >= 1e6) return (v / 1e6).toFixed(1) + 'M';
  if (v >= 1e3) return (v / 1e3).toFixed(0) + 'K';
  return String(Math.round(v));
}

// 종합 실행: 순위 리스트 + 최상위 추천 반환
async function runRecommendation({ universe = UNIVERSE, top = 5 } = {}) {
  const { data, errors } = await fetchAll(universe);

  const scored = [];
  for (const { symbol, bars } of data) {
    const metrics = computeMetrics(bars);
    if (!metrics) continue;
    const s = scoreSymbol(symbol, metrics);
    if (!s || s.filtered) continue;
    s.isETF = ETF_SET.has(symbol);
    scored.push(s);
  }

  scored.sort((a, b) => b.score - a.score);

  return {
    generatedAt: new Date().toISOString(),
    universeSize: universe.length,
    fetched: data.length,
    scoredCount: scored.length,
    errors,
    ranking: scored.slice(0, top),
    pick: scored[0] || null,
  };
}

// 사람이 읽는 텍스트 리포트
function formatReport(result) {
  const L = [];
  const now = new Date(result.generatedAt);
  const kst = now.toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
  L.push('══════════════════════════════════════════════');
  L.push('   📈 오늘의 미국장 단타 추천 종목 (자동 분석)');
  L.push(`   생성: ${kst} (KST)`);
  L.push('══════════════════════════════════════════════');

  const p = result.pick;
  if (!p) {
    L.push('⚠️  조건을 만족하는 추천 종목을 찾지 못했습니다.');
    if (result.errors?.length) {
      L.push(`   (데이터 수집 실패 ${result.errors.length}건 - 네트워크/티커 확인)`);
    }
    return L.join('\n');
  }

  const m = p.metrics;
  L.push('');
  L.push(`🏆 최고 추천: ${p.symbol} ${p.isETF ? '(ETF)' : ''}`);
  L.push(`   종합 점수 : ${p.score} / 100`);
  L.push(`   매매 편향 : ${p.bias}`);
  L.push(`   기준일 종가: $${fmt(m.price)}  (기준봉 ${m.date})`);
  L.push('');
  L.push('   ── 핵심 지표 ──────────────────────');
  L.push(`   변동성 ATR%    : ${fmt(m.atrPct)}%   (장중 기대 변동폭 ≈ $${fmt(m.atr14)})`);
  L.push(`   상대거래량 RVOL: ${fmt(m.rvol)}x  (1.0=평소, 높을수록 관심↑)`);
  L.push(`   5일 모멘텀     : ${fmt(m.mom5)}%`);
  L.push(`   10일 모멘텀    : ${fmt(m.mom10)}%`);
  L.push(`   RSI(14)        : ${fmt(m.rsi14, 1)}`);
  L.push(`   평균 거래대금  : $${humanVol(m.dollarVol)}/일  (유동성)`);
  L.push('');
  L.push('   ── 세부 점수(가중 전) ─────────────');
  L.push(`   변동성 ${fmt(p.parts.volatility, 0)} | 거래량 ${fmt(p.parts.rvol, 0)} | ` +
    `모멘텀 ${fmt(p.parts.momentum, 0)} | 추세 ${fmt(p.parts.trend, 0)} | ` +
    `RSI ${fmt(p.parts.rsi, 0)} | 유동성 ${fmt(p.parts.liquidity, 0)}`);
  L.push('');

  // ATR 기반 리스크 가이드
  const stop = m.atr14 ? m.atr14 * 1.0 : null;
  const target = m.atr14 ? m.atr14 * 1.5 : null;
  L.push('   ── 리스크 가이드(참고) ────────────');
  L.push(`   변동성 기준 손절폭 ≈ $${fmt(stop)} (약 ${fmt(m.atrPct)}%),  ` +
    `1차 목표 ≈ $${fmt(target)} (손익비 1.5)`);
  L.push('');

  if (result.ranking.length > 1) {
    L.push('   ── 상위 후보 순위 ─────────────────');
    result.ranking.forEach((r, i) => {
      L.push(`   ${i + 1}. ${r.symbol.padEnd(6)} ${String(r.score).padStart(5)}점  ` +
        `${r.bias.split('(')[0].padEnd(8)} ATR%=${fmt(r.metrics.atrPct)} RVOL=${fmt(r.metrics.rvol)}`);
    });
    L.push('');
  }

  L.push('──────────────────────────────────────────────');
  L.push('⚠️  본 결과는 과거 일봉 데이터 기반 자동 계산이며 투자 조언이 아닙니다.');
  L.push('    실제 매매는 본인 판단과 리스크 관리 하에 진행하세요.');
  if (result.errors?.length) {
    L.push(`ℹ️  데이터 수집 실패 ${result.errors.length}건 (분석에서 제외됨).`);
  }
  return L.join('\n');
}

// 리포트를 logs/ 에 저장 (txt + 최신 json)
function saveReport(result, reportText, dir = path.join(baseDir(), 'logs')) {
  fs.mkdirSync(dir, { recursive: true });
  const stamp = new Date(result.generatedAt).toISOString().slice(0, 10);
  const txtPath = path.join(dir, `recommendation-${stamp}.txt`);
  const jsonPath = path.join(dir, `recommendation-${stamp}.json`);
  fs.writeFileSync(txtPath, reportText, 'utf8');
  fs.writeFileSync(jsonPath, JSON.stringify(result, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'latest.json'), JSON.stringify(result, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'latest.txt'), reportText, 'utf8');
  return { txtPath, jsonPath };
}

module.exports = { runRecommendation, formatReport, saveReport };
