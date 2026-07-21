// 단타(day-trading) 적합도 스코어링 알고리즘
//
// 아이디어: 장중 단타로 수익을 내려면 아래 조건이 골고루 좋아야 한다.
//   1) 유동성(Liquidity)    - 충분히 거래되어야 진입/청산이 쉽고 슬리피지가 작다  (하드 필터 + 가점)
//   2) 변동성(Volatility)   - 장중 움직임(ATR%)이 커야 먹을 게 있다
//   3) 상대거래량(RVOL)     - 평소보다 거래가 몰리면 촉매/관심 → 장중 방향성 발생 가능성↑
//   4) 모멘텀(Momentum)     - 최근 며칠 추세가 살아있어야 방향성 매매가 쉽다
//   5) 추세 정렬(Trend)     - 가격 > EMA20 > EMA50 이면 롱 편향, 반대면 숏 편향
//   6) RSI 위치            - 극단(과매수/과매도)보다 추세 지속 구간이 단타에 유리
//
// 각 요소를 0~100 으로 정규화 후 가중합. 최종 점수 0~100.

const DEFAULT_WEIGHTS = {
  volatility: 0.28,
  rvol: 0.20,
  momentum: 0.24,
  trend: 0.16,
  rsi: 0.07,
  liquidity: 0.05,
};

// 유동성 하드 필터: 평균 일 거래대금이 이 값 미만이면 후보 제외 (단위: 달러)
const MIN_DOLLAR_VOLUME = 50_000_000; // 5천만 달러

// 값을 [0,100] 으로 클램프
const clamp100 = (x) => Math.max(0, Math.min(100, x));

// 삼각형(최적점) 스코어: x가 sweet 지점에서 100, lo/hi 경계에서 0
function triangular(x, lo, sweet, hi) {
  if (x <= lo || x >= hi) return 0;
  if (x === sweet) return 100;
  if (x < sweet) return ((x - lo) / (sweet - lo)) * 100;
  return ((hi - x) / (hi - sweet)) * 100;
}

function scoreVolatility(atrPct) {
  if (atrPct == null) return 0;
  // 단타 스위트스팟: ATR% ≈ 3.5%. 1% 미만은 너무 잔잔, 9% 초과는 과열/리스크 과다.
  return triangular(atrPct, 0.8, 3.5, 9);
}

function scoreRvol(rvol) {
  if (rvol == null) return 30;
  // 1.0 = 평소 수준(50점), 2.0배 이상이면 만점 근접, 0.5 이하면 저조
  if (rvol <= 0.5) return 10;
  return clamp100(20 + (rvol - 0.5) * 55); // rvol 2.0 → ~102 → 100
}

function scoreMomentum(mom5, mom10) {
  const m5 = mom5 ?? 0;
  const m10 = mom10 ?? 0;
  // 방향 무관하게 "움직임의 강도"를 본다 (롱/숏 양방향 단타 가능)
  const strength = Math.abs(m5) * 0.6 + Math.abs(m10) * 0.4;
  // 약 8% 강도에서 만점 수준
  return clamp100(strength * 12.5);
}

function scoreTrend(m) {
  // 추세가 한 방향으로 깔끔하게 정렬될수록 가점 (롱이든 숏이든)
  let s = 50;
  const bull = m.aboveEma20 && m.aboveEma50 && m.emaStackBull;
  const bear = m.aboveEma20 === false && m.aboveEma50 === false && m.emaStackBull === false;
  if (bull || bear) s = 100;
  else if (m.aboveEma20 === m.aboveEma50) s = 70; // 부분 정렬
  else s = 40; // 혼조
  return s;
}

function scoreRsi(rsi14) {
  if (rsi14 == null) return 50;
  // 추세 지속 구간(45~70 롱, 30~55 숏)을 선호, 극단은 감점
  // 40~65 를 sweet 로 보는 완만한 삼각형
  return triangular(rsi14, 15, 55, 90);
}

function scoreLiquidity(dollarVol) {
  if (dollarVol == null) return 0;
  // 로그 스케일: 5천만 달러=0점 근처, 20억 달러 이상 만점 근처
  const x = Math.log10(dollarVol);
  return clamp100((x - 7.7) * 45); // 10^7.7≈5천만 → 0, 10^9.9≈8B → 100
}

// 하나의 종목 metrics를 받아 세부 점수 + 종합 점수 + 매매 방향 반환
function scoreSymbol(symbol, metrics, weights = DEFAULT_WEIGHTS) {
  if (!metrics) return null;
  // 유동성 하드 필터
  if (metrics.dollarVol != null && metrics.dollarVol < MIN_DOLLAR_VOLUME) {
    return { symbol, filtered: true, reason: '유동성 부족', metrics };
  }

  const parts = {
    volatility: scoreVolatility(metrics.atrPct),
    rvol: scoreRvol(metrics.rvol),
    momentum: scoreMomentum(metrics.mom5, metrics.mom10),
    trend: scoreTrend(metrics),
    rsi: scoreRsi(metrics.rsi14),
    liquidity: scoreLiquidity(metrics.dollarVol),
  };

  const total = Object.entries(weights)
    .reduce((sum, [k, w]) => sum + (parts[k] || 0) * w, 0);

  // 매매 방향 편향 (추세 + 최근 모멘텀 종합)
  const bullVotes = (metrics.aboveEma20 ? 1 : 0) + (metrics.aboveEma50 ? 1 : 0) +
    (metrics.emaStackBull ? 1 : 0) + ((metrics.mom5 ?? 0) > 0 ? 1 : 0) +
    (metrics.closeStrength > 0.55 ? 1 : 0);
  const bias = bullVotes >= 3 ? 'LONG(상승 편향)' : bullVotes <= 1 ? 'SHORT(하락 편향)' : 'NEUTRAL(중립)';

  return {
    symbol,
    filtered: false,
    score: Math.round(total * 10) / 10,
    parts,
    bias,
    metrics,
  };
}

module.exports = {
  scoreSymbol, DEFAULT_WEIGHTS, MIN_DOLLAR_VOLUME,
  scoreVolatility, scoreRvol, scoreMomentum, scoreTrend, scoreRsi, scoreLiquidity,
};
