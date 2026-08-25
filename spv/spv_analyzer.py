#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SPV 파일 분석기 (SPV file analyzer)
====================================

`.SPV` 파일은 산업용 계측/모니터링 장비가 남긴 이진(binary) 시계열 로그입니다.
파일 이름의 `TM095030` 처럼 `TM<시분초>` 형태로 측정 시작 시각이 들어갑니다.

이 스크립트는 아무 외부 라이브러리 없이(순수 파이썬 표준 라이브러리만 사용)
파일 구조를 해석하고 다음을 만들어 줍니다.
  * 콘솔 요약 리포트  (구조 + 채널별 통계)
  * CSV 파일          (표 계산기/엑셀에서 열 수 있음)
  * HTML 리포트       (인라인 SVG 그래프 포함, 브라우저에서 바로 열림)

------------------------------------------------------------------------
역설계로 밝혀낸 파일 포맷 (reverse-engineered format)
------------------------------------------------------------------------
파일은 512바이트 "블록"의 연속입니다. (예: 3584바이트 = 7블록)

각 블록 = 32바이트 헤더 + 480바이트 데이터

[블록 헤더 32바이트]
  off 0  : 4바이트 매직   44 05 1A 08   (모든 블록 동일, 포맷 식별자)
  off 4  : 1바이트        0x18(=24) 고정
  off 5  : 1바이트 시(H)  ── 이 블록 첫 샘플의 시각
  off 6  : 1바이트 분(M)
  off 7  : 1바이트 초(S)
  off 8  : uint16 = 20    샘플 간격(초).  20초마다 한 레코드
  off 10 : uint16        블록 순번 (0,1,2,...)
  off 12 : uint16 = 0     예약
  off 14 : uint16        이 블록에 들어있는 유효 레코드 수 (마지막 블록만 적음)
  off 16 : uint16 = 261   설정값 1 (장비 파라미터로 추정)
  off 18 : uint16 = 1500  설정값 2
  off 20 : int16  = -600  설정값 3
  off 22 : 10바이트 0     예약/패딩

[데이터 레코드 12바이트] × (유효 레코드 수)
  6개의 int16(little-endian) 채널값 = CH1..CH6
    CH1, CH2, CH3, CH5 : 실제로 변하는 측정 신호
    CH4, CH6           : 기준/상태 채널 (측정 중 고정값,
                         종료 구간에서 값이 바뀜 → 상태 플래그로 추정)

각 채널의 물리 단위(전류/전압/힘 등)는 제조사 고유 규격이라 공개되어 있지
않지만, 컨테이너 구조·시간축·채널 데이터는 위와 같이 완전히 복원됩니다.
"""

import struct
import sys
import os
import csv
import argparse
import html

BLOCK_SIZE = 512
HEADER_SIZE = 32
RECORD_SIZE = 12
NUM_CHANNELS = 6
MAGIC = bytes([0x44, 0x05, 0x1A, 0x08])


class SpvBlock:
    def __init__(self, index, raw):
        self.magic = raw[0:4]
        self.hour = raw[5]
        self.minute = raw[6]
        self.second = raw[7]
        self.interval_sec = struct.unpack_from("<H", raw, 8)[0]
        self.block_seq = struct.unpack_from("<H", raw, 10)[0]
        self.record_count = struct.unpack_from("<H", raw, 14)[0]
        self.cfg1 = struct.unpack_from("<H", raw, 16)[0]
        self.cfg2 = struct.unpack_from("<H", raw, 18)[0]
        self.cfg3 = struct.unpack_from("<h", raw, 20)[0]
        self.records = []
        for r in range(self.record_count):
            off = HEADER_SIZE + r * RECORD_SIZE
            self.records.append(struct.unpack_from("<6h", raw, off))

    @property
    def start_time(self):
        return f"{self.hour:02d}:{self.minute:02d}:{self.second:02d}"


class SpvFile:
    def __init__(self, path):
        self.path = path
        with open(path, "rb") as f:
            self.data = f.read()
        self.blocks = []
        self._parse()

    def _parse(self):
        if len(self.data) % BLOCK_SIZE != 0:
            print(f"경고: 파일 크기({len(self.data)})가 512의 배수가 아닙니다. "
                  "마지막 조각은 무시합니다.", file=sys.stderr)
        n_blocks = len(self.data) // BLOCK_SIZE
        for b in range(n_blocks):
            raw = self.data[b * BLOCK_SIZE:(b + 1) * BLOCK_SIZE]
            blk = SpvBlock(b, raw)
            if blk.magic != MAGIC:
                print(f"경고: 블록 {b}의 매직이 예상과 다릅니다 "
                      f"({blk.magic.hex()}).", file=sys.stderr)
            self.blocks.append(blk)

    # ---- 파생 데이터 -------------------------------------------------
    @property
    def interval_sec(self):
        return self.blocks[0].interval_sec if self.blocks else 0

    @property
    def total_records(self):
        return sum(len(b.records) for b in self.blocks)

    def channel(self, idx):
        """0-based 채널 인덱스의 전체 시계열 값 리스트."""
        out = []
        for b in self.blocks:
            for rec in b.records:
                out.append(rec[idx])
        return out

    def timeline(self):
        """각 샘플의 절대 시각(초 단위, 자정 기준)과 'HH:MM:SS' 문자열."""
        times = []
        for b in self.blocks:
            base = b.hour * 3600 + b.minute * 60 + b.second
            for i in range(len(b.records)):
                t = base + i * b.interval_sec
                times.append(t)
        return times

    @staticmethod
    def fmt_hms(total_sec):
        total_sec = int(total_sec)
        return f"{total_sec // 3600:02d}:{(total_sec % 3600) // 60:02d}:{total_sec % 60:02d}"


# ---------------------------------------------------------------------
# 통계
# ---------------------------------------------------------------------
def channel_stats(values):
    n = len(values)
    mn, mx = min(values), max(values)
    mean = sum(values) / n
    var = sum((v - mean) ** 2 for v in values) / n
    distinct = len(set(values))
    return {
        "n": n, "min": mn, "max": mx, "mean": mean,
        "std": var ** 0.5, "range": mx - mn, "distinct": distinct,
        "constant": distinct <= 2,
    }


# ---------------------------------------------------------------------
# 콘솔 리포트
# ---------------------------------------------------------------------
def sparkline(values, width=60):
    chars = " .:-=+*#%@"
    lo, hi = min(values), max(values)
    rng = (hi - lo) or 1
    step = max(1, len(values) // width)
    sampled = values[::step]
    return "".join(chars[int((v - lo) / rng * (len(chars) - 1))] for v in sampled)


def print_report(spv):
    b0 = spv.blocks[0]
    times = spv.timeline()
    print("=" * 66)
    print(" SPV 파일 분석 리포트")
    print("=" * 66)
    print(f" 파일        : {os.path.basename(spv.path)}")
    print(f" 크기        : {len(spv.data):,} 바이트  ({len(spv.blocks)} 블록 × {BLOCK_SIZE}B)")
    print(f" 포맷 매직   : {b0.magic.hex(' ')}  (정상)" if b0.magic == MAGIC
          else f" 포맷 매직   : {b0.magic.hex(' ')}  (예상과 다름!)")
    print(f" 측정 시작   : {b0.start_time}")
    print(f" 측정 종료   : {SpvFile.fmt_hms(times[-1] + spv.interval_sec)}")
    print(f" 총 지속시간 : 약 {SpvFile.fmt_hms(times[-1] - times[0] + spv.interval_sec)}")
    print(f" 샘플 간격   : {spv.interval_sec} 초")
    print(f" 총 샘플수   : {spv.total_records} 개")
    print(f" 채널 수     : {NUM_CHANNELS} 개 (16비트 정수)")
    print(f" 장비 설정   : cfg1={b0.cfg1}, cfg2={b0.cfg2}, cfg3={b0.cfg3}")
    print("-" * 66)
    print(" [블록 구성]")
    print(f"   {'블록':>4} {'시작시각':>10} {'샘플수':>6}")
    for b in spv.blocks:
        print(f"   {b.block_seq:>4} {b.start_time:>10} {b.record_count:>6}")
    print("-" * 66)
    print(" [채널별 통계]")
    print(f"   {'채널':>4} {'최소':>7} {'최대':>7} {'평균':>9} {'표준편차':>9} {'구분':>10}")
    for c in range(NUM_CHANNELS):
        st = channel_stats(spv.channel(c))
        kind = "기준/상태" if st["constant"] else "측정신호"
        print(f"   CH{c+1:>2} {st['min']:>7} {st['max']:>7} "
              f"{st['mean']:>9.1f} {st['std']:>9.1f} {kind:>10}")
    print("-" * 66)
    print(" [채널 파형 미리보기]  낮음' '←→'@'높음")
    for c in range(NUM_CHANNELS):
        print(f"   CH{c+1}: |{sparkline(spv.channel(c))}|")
    print("=" * 66)


# ---------------------------------------------------------------------
# CSV 내보내기
# ---------------------------------------------------------------------
def export_csv(spv, path):
    times = spv.timeline()
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["index", "time", "seconds"] + [f"CH{c+1}" for c in range(NUM_CHANNELS)])
        chans = [spv.channel(c) for c in range(NUM_CHANNELS)]
        for i, t in enumerate(times):
            w.writerow([i, SpvFile.fmt_hms(t), t] + [chans[c][i] for c in range(NUM_CHANNELS)])
    return path


# ---------------------------------------------------------------------
# HTML 리포트 (인라인 SVG 그래프)
# ---------------------------------------------------------------------
def svg_line_chart(values, times, title, w=760, h=180, color="#2563eb"):
    pad_l, pad_r, pad_t, pad_b = 54, 12, 24, 22
    iw, ih = w - pad_l - pad_r, h - pad_t - pad_b
    lo, hi = min(values), max(values)
    rng = (hi - lo) or 1
    n = len(values)

    def px(i):
        return pad_l + (i / (n - 1 if n > 1 else 1)) * iw

    def py(v):
        return pad_t + ih - (v - lo) / rng * ih

    pts = " ".join(f"{px(i):.1f},{py(v):.1f}" for i, v in enumerate(values))
    # y축 눈금 3개
    yticks = ""
    for k in range(3):
        val = lo + rng * k / 2
        y = py(val)
        yticks += (f'<line x1="{pad_l}" y1="{y:.1f}" x2="{w-pad_r}" y2="{y:.1f}" '
                   f'stroke="#e5e7eb" stroke-width="1"/>'
                   f'<text x="{pad_l-6}" y="{y+3:.1f}" text-anchor="end" '
                   f'font-size="10" fill="#6b7280">{val:.0f}</text>')
    # x축 라벨 (시작/중간/끝)
    xlabels = ""
    for frac in (0.0, 0.5, 1.0):
        i = int((n - 1) * frac)
        xlabels += (f'<text x="{px(i):.1f}" y="{h-6}" text-anchor="middle" '
                    f'font-size="10" fill="#6b7280">{SpvFile.fmt_hms(times[i])}</text>')
    return f"""<svg viewBox="0 0 {w} {h}" width="100%" style="max-width:{w}px">
  <text x="{pad_l}" y="14" font-size="12" font-weight="600" fill="#111827">{html.escape(title)}</text>
  {yticks}
  <polyline fill="none" stroke="{color}" stroke-width="1.6" points="{pts}"/>
  {xlabels}
</svg>"""


def export_html(spv, path):
    times = spv.timeline()
    b0 = spv.blocks[0]
    colors = ["#2563eb", "#16a34a", "#db2777", "#9333ea", "#ea580c", "#0891b2"]
    charts = ""
    for c in range(NUM_CHANNELS):
        st = channel_stats(spv.channel(c))
        kind = "기준/상태 채널" if st["constant"] else "측정 신호"
        charts += f"""<div class="card">
  <div class="chead">CH{c+1} <span class="tag">{kind}</span>
    <span class="stat">min {st['min']} · max {st['max']} · 평균 {st['mean']:.0f}</span></div>
  {svg_line_chart(spv.channel(c), times, f"CH{c+1}", color=colors[c])}
</div>"""

    block_rows = "".join(
        f"<tr><td>{b.block_seq}</td><td>{b.start_time}</td><td>{b.record_count}</td></tr>"
        for b in spv.blocks)

    doc = f"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SPV 분석 리포트</title>
<style>
  body{{font-family:system-ui,'Segoe UI',sans-serif;margin:0;background:#f8fafc;color:#0f172a}}
  .wrap{{max-width:860px;margin:0 auto;padding:24px}}
  h1{{font-size:22px;margin:0 0 4px}}
  .sub{{color:#64748b;font-size:13px;margin-bottom:20px}}
  .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:24px}}
  .kv{{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px 14px}}
  .kv b{{display:block;font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.03em}}
  .kv span{{font-size:18px;font-weight:700}}
  .card{{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px 14px;margin-bottom:14px}}
  .chead{{font-weight:700;font-size:14px;margin-bottom:4px}}
  .tag{{font-size:11px;background:#eef2ff;color:#4338ca;border-radius:6px;padding:2px 6px;margin-left:4px}}
  .stat{{float:right;font-weight:400;font-size:12px;color:#64748b}}
  table{{border-collapse:collapse;width:100%;font-size:13px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden}}
  th,td{{padding:6px 10px;text-align:left;border-bottom:1px solid #f1f5f9}}
  th{{background:#f8fafc;color:#475569;font-size:12px}}
  h2{{font-size:15px;margin:22px 0 10px}}
  .note{{background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:12px 14px;font-size:13px;color:#713f12;line-height:1.6}}
</style></head><body><div class="wrap">
<h1>SPV 파일 분석 리포트</h1>
<div class="sub">{html.escape(os.path.basename(spv.path))} · {len(spv.data):,} 바이트</div>

<div class="grid">
  <div class="kv"><b>측정 시작</b><span>{b0.start_time}</span></div>
  <div class="kv"><b>측정 종료</b><span>{SpvFile.fmt_hms(times[-1]+spv.interval_sec)}</span></div>
  <div class="kv"><b>샘플 간격</b><span>{spv.interval_sec}초</span></div>
  <div class="kv"><b>총 샘플</b><span>{spv.total_records}</span></div>
  <div class="kv"><b>채널</b><span>{NUM_CHANNELS}</span></div>
  <div class="kv"><b>블록</b><span>{len(spv.blocks)}</span></div>
</div>

<h2>채널 파형</h2>
{charts}

<h2>블록 구성</h2>
<table><tr><th>블록</th><th>시작 시각</th><th>샘플 수</th></tr>{block_rows}</table>

<h2>포맷 설명</h2>
<div class="note">
이 <b>.SPV</b> 파일은 산업용 계측 장비의 이진 시계열 로그입니다.
512바이트 블록의 연속으로, 각 블록은 32바이트 헤더(매직 <code>44 05 1A 08</code>,
블록 시작 시각, 샘플 간격 {spv.interval_sec}초, 장비 설정 cfg1={b0.cfg1}/cfg2={b0.cfg2}/cfg3={b0.cfg3})와
{spv.interval_sec}초 간격으로 기록된 12바이트 레코드(6개 16비트 채널)들로 구성됩니다.
CH1·CH2·CH3·CH5는 실제 변화하는 측정 신호이고, CH4·CH6은 값이 거의 고정된 기준/상태 채널입니다.
각 채널의 물리 단위는 제조사 고유 규격이라 파일 자체에는 들어있지 않습니다.
</div>
</div></body></html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    return path


# ---------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="SPV 이진 로그 파일 분석기")
    ap.add_argument("file", help="분석할 .SPV 파일 경로")
    ap.add_argument("--csv", metavar="PATH", nargs="?", const="__auto__",
                    help="CSV로 내보내기 (경로 생략 시 입력파일명.csv)")
    ap.add_argument("--html", metavar="PATH", nargs="?", const="__auto__",
                    help="HTML 리포트 생성 (경로 생략 시 입력파일명.html)")
    ap.add_argument("-q", "--quiet", action="store_true", help="콘솔 리포트 생략")
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        print(f"오류: 파일을 찾을 수 없습니다: {args.file}", file=sys.stderr)
        sys.exit(1)

    spv = SpvFile(args.file)
    if not spv.blocks:
        print("오류: 유효한 블록이 없습니다.", file=sys.stderr)
        sys.exit(1)

    if not args.quiet:
        print_report(spv)

    stem = os.path.splitext(args.file)[0]
    if args.csv is not None:
        out = stem + ".csv" if args.csv == "__auto__" else args.csv
        print(f"\nCSV 저장됨 → {export_csv(spv, out)}")
    if args.html is not None:
        out = stem + ".html" if args.html == "__auto__" else args.html
        print(f"HTML 저장됨 → {export_html(spv, out)}")


if __name__ == "__main__":
    main()
