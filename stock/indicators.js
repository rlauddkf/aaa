// 기술적 지표 계산 (외부 라이브러리 없이 순수 구현)
// 모든 함수는 일봉 배열(과거 -> 최신 순)을 입력받습니다.

function sma(values, period) {
  if (values.length < period) return null;
  let sum = 0;
  for (let i = values.length - period; i < values.length; i++) sum += values[i];
  return sum / period;
}

function ema(values, period) {
  if (values.length < period) return null;
  const k = 2 / (period + 1);
  // 초기값: 첫 period 구간의 SMA
  let prev = values.slice(0, period).reduce((a, b) => a + b, 0) / period;
  for (let i = period; i < values.length; i++) {
    prev = values[i] * k + prev * (1 - k);
  }
  return prev;
}

// RSI (Wilder 방식)
function rsi(closes, period = 14) {
  if (closes.length < period + 1) return null;
  let gain = 0, loss = 0;
  for (let i = 1; i <= period; i++) {
    const diff = closes[i] - closes[i - 1];
    if (diff >= 0) gain += diff; else loss -= diff;
  }
  let avgGain = gain / period, avgLoss = loss / period;
  for (let i = period + 1; i < closes.length; i++) {
    const diff = closes[i] - closes[i - 1];
    const g = diff > 0 ? diff : 0;
    const l = diff < 0 ? -diff : 0;
    avgGain = (avgGain * (period - 1) + g) / period;
    avgLoss = (avgLoss * (period - 1) + l) / period;
  }
  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - 100 / (1 + rs);
}

// ATR (Average True Range) - 절대값
function atr(bars, period = 14) {
  if (bars.length < period + 1) return null;
  const trs = [];
  for (let i = 1; i < bars.length; i++) {
    const h = bars[i].high, l = bars[i].low, pc = bars[i - 1].close;
    trs.push(Math.max(h - l, Math.abs(h - pc), Math.abs(l - pc)));
  }
  // 최근 period개 TR의 평균
  const recent = trs.slice(-period);
  return recent.reduce((a, b) => a + b, 0) / recent.length;
}

// 전 종목 지표를 한 번에 계산
function computeMetrics(bars) {
  if (!bars || bars.length < 20) return null;
  const closes = bars.map((b) => b.close);
  const volumes = bars.map((b) => b.volume);
  const last = bars[bars.length - 1];
  const price = last.close;

  const atr14 = atr(bars, 14);
  const rsi14 = rsi(closes, 14);
  const ema20 = ema(closes, 20);
  const ema50 = ema(closes, Math.min(50, closes.length - 1));
  const sma20vol = sma(volumes, 20);

  // 변동성: ATR을 가격 대비 % 로
  const atrPct = atr14 != null ? (atr14 / price) * 100 : null;

  // 상대 거래량 (오늘 거래량 / 20일 평균 거래량)
  const rvol = sma20vol ? last.volume / sma20vol : null;

  // 평균 일 거래대금 (유동성) = 20일 평균거래량 * 현재가
  const dollarVol = sma20vol ? sma20vol * price : null;

  // 모멘텀
  const mom5 = closes.length > 5 ? (price / closes[closes.length - 6] - 1) * 100 : null;
  const mom10 = closes.length > 10 ? (price / closes[closes.length - 11] - 1) * 100 : null;

  // 추세: 현재가가 EMA20/EMA50 위에 있는지
  const aboveEma20 = ema20 != null ? price > ema20 : null;
  const aboveEma50 = ema50 != null ? price > ema50 : null;
  const emaStackBull = ema20 != null && ema50 != null ? ema20 > ema50 : null;

  // 오늘 봉의 강도 (종가가 당일 고저 범위의 어디에 위치하는가)
  const range = last.high - last.low;
  const closeStrength = range > 0 ? (last.close - last.low) / range : 0.5;

  return {
    price, date: last.date,
    atr14, atrPct, rsi14, ema20, ema50,
    rvol, dollarVol, mom5, mom10,
    aboveEma20, aboveEma50, emaStackBull, closeStrength,
    volume: last.volume,
  };
}

module.exports = { sma, ema, rsi, atr, computeMetrics };
