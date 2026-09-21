# -*- coding: utf-8 -*-
"""
latency_toolkit —— 字幕/转录显示时延自动测量工具（可分享版）

流程：
  1) asr      对视频做语音识别，得到词级时间戳 words.json
  2) ocrscan  按低帧率（默认 5fps）扫视频，对每侧手机屏幕 ROI 做 OCR，得到 frames.jsonl
  3) calc     结合 句子文本 + words.json + frames.jsonl，计算每句 语音起止 / 各侧显示起止 / 时延，
              可选 --refine 用原视频逐帧 OCR 精修到 60fps，并可选与剪映人工标记对比
  4) marks    单独读取剪映复合片段里的人工标记，导出 Excel（不依赖 1-3 步）
  5) all      一次性跑 asr + ocrscan + calc

快速开始：
  pip install -r requirements.txt
  python cli.py all --video "D:/my.mp4" --sentences sentences.txt \
      --roi L=300,880 --roi R=920,1520 --y0 400 --y1 960 --out result.xlsx

sentences.txt：每行一句台词原文（英文/中文均可，用于跟 ASR 结果对齐）。
--roi 可重复传入多次，key=x0,x1 的形式；每个 key 就是一侧（比如 L / R / 手机A / 手机B）。
"""
import argparse
import json
import statistics
import sys
import time
from pathlib import Path

DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(DIR))

import core
import marks as marks_mod

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def parse_roi(items):
    roi = {}
    for it in items or []:
        if "=" not in it:
            raise SystemExit(f"--roi 格式应为 key=x0,x1，收到: {it}")
        k, v = it.split("=", 1)
        x0, x1 = v.split(",")
        roi[k.strip()] = (int(x0), int(x1))
    return roi


# ---------------------------------------------------------------- asr

def cmd_asr(a):
    from faster_whisper import WhisperModel
    t0 = time.time()
    model = WhisperModel(a.model, device=a.device, compute_type=a.compute_type)
    segs, info = model.transcribe(a.video, language=a.lang, word_timestamps=True,
                                   beam_size=5, vad_filter=False)
    words = []
    for s in segs:
        for w in s.words:
            words.append({"w": w.word.strip(), "s": round(w.start, 3), "e": round(w.end, 3),
                          "p": round(w.probability, 2)})
    json.dump(words, open(a.out, "w", encoding="utf8"), ensure_ascii=False, indent=0)
    print(f"{len(words)} 个词，用时 {time.time()-t0:.1f}s -> {a.out}")


# ---------------------------------------------------------------- ocrscan

def cmd_ocrscan(a):
    import cv2
    from rapidocr_onnxruntime import RapidOCR
    roi = parse_roi(a.roi)
    if not roi:
        raise SystemExit("请用 --roi key=x0,x1 指定至少一个 ROI")
    ocr = RapidOCR()
    cap = cv2.VideoCapture(a.video)
    fps = cap.get(cv2.CAP_PROP_FPS)
    t_end = a.t_end
    if t_end is None:
        t_end = cap.get(cv2.CAP_PROP_FRAME_COUNT) / fps
    out = open(a.out, "w", encoding="utf8")
    t = a.t_start
    n = 0
    t0 = time.time()
    while t <= t_end:
        cap.set(cv2.CAP_PROP_POS_FRAMES, round(t * fps))
        ok, fr = cap.read()
        if not ok:
            break
        rec = {"t": round(t, 2)}
        for side, (x0, x1) in roi.items():
            res, _ = ocr(fr[a.y0:a.y1, x0:x1])
            rec[side] = [[round(b[0][1]) + a.y0, txt, round(sc, 2)] for b, txt, sc in (res or [])
                         if any(c.isascii() and c.isalpha() for c in txt)]
        out.write(json.dumps(rec, ensure_ascii=False) + "\n")
        out.flush()
        n += 1
        if n % 50 == 0:
            print(f"t={t:.1f} frames={n} elapsed={time.time()-t0:.0f}s", flush=True)
        t += a.step
    out.close()
    print(f"完成，{n} 帧，用时 {time.time()-t0:.0f}s -> {a.out}")


# ---------------------------------------------------------------- calc

def fmt(x):
    return "" if x is None else round(x, 3)


def build_excel(sents, speech, shows, side_keys, mrows, mark_a, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "自动时延"
    hdr = ["序号", "文本", "语音起", "语音止"]
    for sk in side_keys:
        hdr += [f"{sk} 显示起", f"{sk} 显示止"]
    for sk in side_keys:
        hdr += [f"句首时延 {sk}", f"句尾时延 {sk}"]
    hdr.append("对应人工标记(S)")
    ws.append(hdr)
    bold = Font(bold=True); center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ncol = len(hdr)
    for c in range(1, ncol + 1):
        ws.cell(1, c).font = bold; ws.cell(1, c).alignment = center
        ws.cell(1, c).fill = PatternFill("solid", fgColor="E8E8E8")

    base_col = 5 + 2 * len(side_keys)  # 时延列起始
    FIRST = 2
    for i, s in enumerate(sents):
        r = FIRST + i
        sp = speech[i]
        ws.cell(r, 1, i + 1); ws.cell(r, 2, s)
        if sp:
            ws.cell(r, 3, round(sp["start"], 3)); ws.cell(r, 4, round(sp["end"], 3))
        c = 5
        for sk in side_keys:
            v = shows[sk][i]
            if v["start"] is not None:
                ws.cell(r, c, v["start"])
            if v["end"] is not None:
                ws.cell(r, c + 1, v["end"])
            c += 2
        c = base_col
        for j, sk in enumerate(side_keys):
            start_col = get_column_letter(5 + 2 * j)
            end_col = get_column_letter(6 + 2 * j)
            ws.cell(r, c, f'=IF(AND({start_col}{r}<>"",C{r}<>""),{start_col}{r}-C{r},"")')
            ws.cell(r, c + 1, f'=IF(AND({end_col}{r}<>"",D{r}<>""),{end_col}{r}-D{r},"")')
            c += 2
        if mrows:
            ms = [n for sec, n in mark_a if sp and sp["start"] - 0.3 <= sec <= sp["end"] - 0.3]
            ws.cell(r, ncol, ",".join(f"S{m}" for m in ms))
        for cc in range(3, ncol):
            ws.cell(r, cc).number_format = "0.000"; ws.cell(r, cc).alignment = Alignment(horizontal="center")
    LAST = FIRST + len(sents) - 1
    r = LAST + 1
    ws.cell(r, 2, "平均").font = bold
    for c in range(base_col, ncol):
        L = get_column_letter(c)
        ws.cell(r, c, f'=IFERROR(AVERAGE({L}{FIRST}:{L}{LAST}),"")'); ws.cell(r, c).number_format = "0.000"
    widths = {1: 5, 2: 60}
    for c in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(c, 10)
    ws.freeze_panes = "C2"

    ws2 = wb.create_sheet("匹配细节")
    ws2.append(["序号"] + [f"{sk} 首词" for sk in side_keys] + [f"{sk} 尾词" for sk in side_keys] + ["ASR 文本", "区分词"])
    for i in range(len(sents)):
        row = [i + 1] + [shows[sk][i].get("head") for sk in side_keys] + [shows[sk][i].get("tail") for sk in side_keys]
        row += [speech[i]["asr"] if speech[i] else "", shows[side_keys[0]][i].get("dist", "")]
        ws2.append(row)
    try:
        wb.save(out_path)
    except PermissionError:
        out_path = out_path.with_name(out_path.stem + time.strftime("_%H%M%S") + ".xlsx")
        wb.save(out_path)
    return out_path


def cmd_calc(a):
    roi = parse_roi(a.roi) if a.roi else {}
    side_keys = list(roi.keys()) if roi else None
    sents = [l.strip() for l in open(a.sentences, encoding="utf8") if l.strip()]
    words = json.load(open(a.words, encoding="utf8"))
    frames = core.load_frames(a.frames)
    if side_keys is None:
        side_keys = sorted({k for fr in frames for k in fr if k != "t"})
    print(f"句子 {len(sents)}，ASR 词 {len(words)}，OCR 帧 {len(frames)}（{frames[0]['t']}s–{frames[-1]['t']}s），侧: {side_keys}")

    speech = core.speech_times(sents, words)
    wtimes = core.tok_times(words, sents)
    shows = {sk: core.show_times(sents, speech, frames, sk, wtimes) for sk in side_keys}

    if a.refine:
        if not roi or not a.video:
            raise SystemExit("--refine 需要同时提供 --video 和 --roi（用于逐帧重新 OCR）")
        ocr = core.OcrEngine(a.video, roi, a.y0, a.y1, cache_path=a.cache)
        t0 = time.time()
        for sk in side_keys:
            core.refine_all(sents, speech, frames, shows[sk], sk, ocr)
        print(f"精修完成，用时 {time.time()-t0:.0f}s")

    mrows, mark_a = {}, []
    if a.draft:
        cands = marks_mod.find_candidates(a.draft, extra_root=a.drafts_root)
        if cands:
            mrows, _, _ = marks_mod.parse_items(cands[0]["items"])
            mark_a = sorted((mrows[n]["a"][0], n) for n in mrows if "a" in mrows[n])
            print(f"已读取人工标记来源: [{cands[0]['draft']}]")
        else:
            print("未找到匹配的剪映人工标记（跳过对比）")

    print(f"{'句':<3}{'语音起':>8}{'语音止':>8}", end="")
    for sk in side_keys:
        print(f"{sk+'起':>8}{sk+'止':>8}", end="")
    print(" | 时延(起/止 每侧)")
    diffs = {sk: [] for sk in side_keys}
    for i, s in enumerate(sents):
        sp = speech[i]
        if not sp:
            print(f"{i+1:<3} 未对齐")
            continue
        f = lambda v: f"{v:8.2f}" if v is not None else f"{'—':>8}"
        print(f"{i+1:<3}{sp['start']:8.2f}{sp['end']:8.2f}", end="")
        line = []
        for sk in side_keys:
            v = shows[sk][i]
            print(f"{f(v['start'])}{f(v['end'])}", end="")
            if v["start"] is not None:
                diffs[sk].append(v["start"] - sp["start"])
            line.append(sk)
        print(" | " + s[:40])
    for sk in side_keys:
        if diffs[sk]:
            print(f"句首时延 {sk}: 均值 {statistics.mean(diffs[sk]):+.3f}  中位 {statistics.median(diffs[sk]):+.3f}  n={len(diffs[sk])}")

    out = build_excel(sents, speech, shows, side_keys, mrows, mark_a, Path(a.out))
    print("saved:", out)


# ---------------------------------------------------------------- marks

def cmd_marks(a):
    cands = marks_mod.find_candidates(a.draft, extra_root=a.drafts_root)
    if not cands:
        print("未找到含标记的明文复合片段文件。请在剪映中：选中片段 → 右键「新建复合片段」→ Ctrl+S 保存后重试。")
        sys.exit(2)
    if a.list:
        for c in cands:
            print(f"{time.strftime('%m-%d %H:%M:%S', time.localtime(c['mtime']))}  {c['n']:3d} 个标记  [{c['draft']}]  {c['file']}")
        return
    src = cands[0]
    rows, anomalies, ignored = marks_mod.parse_items(src["items"])
    out = marks_mod.build_excel(rows, anomalies, ignored, f"[{src['draft']}] {src['file']}", Path(a.out))
    print(f"来源: [{src['draft']}]  文件 {src['file']}")
    for gk in ("start", "end"):
        marks_mod.print_group(marks_mod.GROUPS[gk], rows)
    print("\n检查提示:" if anomalies else "\n检查提示: 无异常")
    for x in anomalies:
        print("  -", x)
    print("\nsaved:", out)


# ---------------------------------------------------------------- all

def cmd_all(a):
    words_path = Path(a.words) if a.words else Path(a.out).with_suffix(".words.json")
    frames_path = Path(a.frames) if a.frames else Path(a.out).with_suffix(".frames.jsonl")
    if not words_path.exists() or a.force:
        cmd_asr(argparse.Namespace(video=a.video, model=a.model, device=a.device,
                                    compute_type=a.compute_type, lang=a.lang, out=str(words_path)))
    else:
        print(f"复用已存在: {words_path}")
    if not frames_path.exists() or a.force:
        cmd_ocrscan(argparse.Namespace(video=a.video, roi=a.roi, y0=a.y0, y1=a.y1,
                                        step=a.step, t_start=a.t_start, t_end=a.t_end, out=str(frames_path)))
    else:
        print(f"复用已存在: {frames_path}")
    cmd_calc(argparse.Namespace(sentences=a.sentences, words=str(words_path), frames=str(frames_path),
                                 roi=a.roi, refine=a.refine, video=a.video, y0=a.y0, y1=a.y1,
                                 cache=a.cache, draft=a.draft, drafts_root=a.drafts_root, out=a.out))


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("asr", help="语音识别，生成词级时间戳 words.json")
    p.add_argument("--video", required=True)
    p.add_argument("--model", default="small", help="faster-whisper 模型名，如 small/medium/large-v3")
    p.add_argument("--device", default="cpu")
    p.add_argument("--compute-type", default="int8")
    p.add_argument("--lang", default="en")
    p.add_argument("--out", default="words.json")
    p.set_defaults(func=cmd_asr)

    p = sub.add_parser("ocrscan", help="低帧率扫描视频 ROI 做 OCR，生成 frames.jsonl")
    p.add_argument("--video", required=True)
    p.add_argument("--roi", action="append", required=True, help="key=x0,x1，可重复")
    p.add_argument("--y0", type=int, required=True)
    p.add_argument("--y1", type=int, required=True)
    p.add_argument("--step", type=float, default=0.2, help="采样间隔秒，默认 0.2（5fps）")
    p.add_argument("--t-start", type=float, default=0.0)
    p.add_argument("--t-end", type=float, default=None)
    p.add_argument("--out", default="frames.jsonl")
    p.set_defaults(func=cmd_ocrscan)

    p = sub.add_parser("calc", help="计算每句语音/显示时延，导出 Excel")
    p.add_argument("--sentences", required=True, help="每行一句台词的文本文件")
    p.add_argument("--words", default="words.json")
    p.add_argument("--frames", default="frames.jsonl")
    p.add_argument("--roi", action="append", help="key=x0,x1，用于 --refine 精修（可选）")
    p.add_argument("--refine", action="store_true", help="用原视频逐帧 OCR 精修到 60fps（较慢）")
    p.add_argument("--video", help="--refine 时需要")
    p.add_argument("--y0", type=int, default=0)
    p.add_argument("--y1", type=int, default=0)
    p.add_argument("--cache", default="ocr_cache.json", help="逐帧 OCR 结果缓存文件")
    p.add_argument("--draft", help="剪映草稿名（用于与人工标记对比，可选）")
    p.add_argument("--drafts-root", help="剪映草稿根目录（可选，自动检测失败时指定）")
    p.add_argument("--out", default="latency_result.xlsx")
    p.set_defaults(func=cmd_calc)

    p = sub.add_parser("marks", help="读取剪映人工标记并导出 Excel（独立功能，不依赖 asr/ocrscan）")
    p.add_argument("--draft", help="草稿名（包含匹配）")
    p.add_argument("--drafts-root", help="剪映草稿根目录（可选）")
    p.add_argument("--list", action="store_true")
    p.add_argument("--out", default="剪映标记时间差.xlsx")
    p.set_defaults(func=cmd_marks)

    p = sub.add_parser("all", help="一次性跑 asr + ocrscan + calc")
    p.add_argument("--video", required=True)
    p.add_argument("--sentences", required=True)
    p.add_argument("--roi", action="append", required=True, help="key=x0,x1，可重复")
    p.add_argument("--y0", type=int, required=True)
    p.add_argument("--y1", type=int, required=True)
    p.add_argument("--step", type=float, default=0.2)
    p.add_argument("--t-start", type=float, default=0.0)
    p.add_argument("--t-end", type=float, default=None)
    p.add_argument("--model", default="small")
    p.add_argument("--device", default="cpu")
    p.add_argument("--compute-type", default="int8")
    p.add_argument("--lang", default="en")
    p.add_argument("--refine", action="store_true")
    p.add_argument("--cache", default="ocr_cache.json")
    p.add_argument("--draft", help="剪映草稿名（可选，用于对比人工标记）")
    p.add_argument("--drafts-root", help="剪映草稿根目录（可选）")
    p.add_argument("--words", help="复用已有 words.json（跳过 asr）")
    p.add_argument("--frames", help="复用已有 frames.jsonl（跳过 ocrscan）")
    p.add_argument("--force", action="store_true", help="即使已存在也重新跑 asr/ocrscan")
    p.add_argument("--out", default="latency_result.xlsx")
    p.set_defaults(func=cmd_all)

    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
