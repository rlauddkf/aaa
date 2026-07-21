// 일봉(OHLCV) 데이터 수집 모듈
// 무료 공개 데이터 소스만 사용 (API 키 불필요)
//  1순위: Yahoo Finance chart API
//  2순위: Stooq CSV (Yahoo 실패 시 폴백)
//
// 반환 형식: { symbol, bars: [{date, open, high, low, close, volume}, ...] }  (과거 -> 최신 순)

const YAHOO_HOSTS = [
  'https://query1.finance.yahoo.com',
  'https://query2.finance.yahoo.com',
];

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
  '(KHTML, like Gecko) Chrome/122.0 Safari/537.36';

// 간단한 sleep
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function fetchYahoo(symbol, { range = '3mo', interval = '1d' } = {}) {
  let lastErr;
  for (const host of YAHOO_HOSTS) {
    const url = `${host}/v8/finance/chart/${encodeURIComponent(symbol)}` +
      `?range=${range}&interval=${interval}&includePrePost=false`;
    try {
      const res = await fetch(url, { headers: { 'User-Agent': UA } });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      const result = json?.chart?.result?.[0];
      if (!result) throw new Error('빈 응답');
      const ts = result.timestamp || [];
      const q = result.indicators?.quote?.[0] || {};
      const bars = [];
      for (let i = 0; i < ts.length; i++) {
        const o = q.open?.[i], h = q.high?.[i], l = q.low?.[i];
        const c = q.close?.[i], v = q.volume?.[i];
        if ([o, h, l, c].some((x) => x == null)) continue;
        bars.push({
          date: new Date(ts[i] * 1000).toISOString().slice(0, 10),
          open: o, high: h, low: l, close: c, volume: v || 0,
        });
      }
      if (bars.length) return { symbol, bars, source: 'yahoo' };
      throw new Error('유효 봉 없음');
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr || new Error('Yahoo 실패');
}

async function fetchStooq(symbol) {
  // Stooq는 미국 종목에 .us 접미사 사용
  const s = symbol.toLowerCase().replace(/\.us$/, '') + '.us';
  const url = `https://stooq.com/q/d/l/?s=${s}&i=d`;
  const res = await fetch(url, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const text = await res.text();
  const lines = text.trim().split('\n');
  if (lines.length < 2 || !/^Date,/i.test(lines[0])) throw new Error('Stooq 데이터 없음');
  const bars = [];
  for (let i = 1; i < lines.length; i++) {
    const [date, o, h, l, c, v] = lines[i].split(',');
    if (!date || o === 'N/D') continue;
    bars.push({
      date, open: +o, high: +h, low: +l, close: +c, volume: +v || 0,
    });
  }
  // 최근 ~70거래일만 사용
  return { symbol, bars: bars.slice(-70), source: 'stooq' };
}

async function fetchBars(symbol, opts) {
  try {
    return await fetchYahoo(symbol, opts);
  } catch (e1) {
    try {
      return await fetchStooq(symbol);
    } catch (e2) {
      throw new Error(`${symbol} 수집 실패 (yahoo: ${e1.message}, stooq: ${e2.message})`);
    }
  }
}

// 여러 심볼을 동시성 제한하며 수집
async function fetchAll(symbols, { concurrency = 6, opts } = {}) {
  const out = [];
  const errors = [];
  let idx = 0;
  async function worker() {
    while (idx < symbols.length) {
      const sym = symbols[idx++];
      try {
        out.push(await fetchBars(sym, opts));
      } catch (e) {
        errors.push({ symbol: sym, error: e.message });
      }
      await sleep(120); // rate-limit 완화
    }
  }
  const workers = Array.from({ length: Math.min(concurrency, symbols.length) }, worker);
  await Promise.all(workers);
  return { data: out, errors };
}

module.exports = { fetchBars, fetchAll, fetchYahoo, fetchStooq };
