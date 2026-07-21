// 미국 시장 단타(day-trading) 후보 유니버스
// - 유동성이 풍부하고 거래량/변동성이 커서 장중 단타에 적합한 종목/ETF 위주로 구성
// - 원하는 종목을 자유롭게 추가/삭제하세요. 티커는 Yahoo Finance 기준입니다.

// 개별 종목 (대형 + 변동성 있는 인기 단타 종목)
const STOCKS = [
  'AAPL', 'MSFT', 'NVDA', 'AMD', 'TSLA', 'AMZN', 'META', 'GOOGL', 'NFLX',
  'AVGO', 'MU', 'INTC', 'QCOM', 'SMCI', 'PLTR', 'COIN', 'MARA', 'RIOT',
  'CRM', 'ORCL', 'ADBE', 'UBER', 'SHOP', 'SNOW', 'ARM', 'DELL', 'ON',
  'BABA', 'PDD', 'JPM', 'BAC', 'DIS', 'BA', 'CAT', 'XOM', 'CVX',
  'LLY', 'UNH', 'PFE', 'MRNA', 'SOFI', 'AFRM', 'RIVN', 'LCID', 'F', 'GM',
];

// ETF (지수/섹터/레버리지 - 단타에 자주 쓰이는 것들)
const ETFS = [
  'SPY', 'QQQ', 'IWM', 'DIA',            // 지수
  'TQQQ', 'SQQQ', 'SOXL', 'SOXS',        // 레버리지 (반도체/나스닥)
  'SPXL', 'SPXS', 'UPRO', 'SDOW',        // 레버리지 (S&P/다우)
  'XLF', 'XLE', 'XLK', 'XLV', 'XLI',     // 섹터
  'SMH', 'ARKK', 'TNA', 'TZA',           // 반도체/혁신/러셀 레버리지
  'GLD', 'SLV', 'USO', 'UNG',            // 원자재
  'FNGU', 'BULZ', 'LABU',                // 빅테크/바이오 레버리지
];

const UNIVERSE = [...STOCKS, ...ETFS];

module.exports = { UNIVERSE, STOCKS, ETFS };
