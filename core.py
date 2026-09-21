# -*- coding: utf-8 -*-
"""
字幕/转录时延测量 —— 核心算法模块
  音频侧：faster-whisper 词级时间戳 + 按句子文本对齐 → 每句 speech_start / speech_end
  显示侧：视频按固定帧率采样 + RapidOCR（每侧手机一个 ROI）→ 每句首/尾"区分词"首次出现的时刻
  精修：命中帧附近用原视频按 60fps（或指定帧率）逐帧 OCR 回溯，找到文字第一次可辨认的时刻，
        再回溯气泡滑入动画的起点（灰度帧差）。
本文件不含任何机器专属路径，所有输入均由调用方传入。
"""
import difflib
import json
import re
from pathlib import Path

from rapidfuzz import fuzz

# ---------------------------------------------------------------- 文本处理

STOP = set(
    "the a an of to and in on at for with that this is are was were be been it its i you he she "
    "they we his her their our my me him them as by from or but not so if do did does have has "
    "had will would can could just very also".split()
)


def norm(t):
    return re.sub(r"[^a-z0-9' ]", " ", t.lower().replace("%", " percent ")).strip()


def toks(t):
    return [x for x in re.split(r"[\s\-]+", norm(t)) if x]


def tok_present(tok, txt):
    """判断 tok 是否已在 txt（OCR 归一化文本）中出现；短词用词边界正则，长词用模糊匹配以容忍 OCR 误识别。"""
    if len(tok) >= 5:
        return fuzz.partial_ratio(tok, txt) >= 90
    return re.search(r"(?<![a-z])" + re.escape(tok) + r"(?![a-z])", txt) is not None


def distinctive(sents, i, span=2):
    """第 i 句里，相邻 span 句都不包含的词（用于避免把上一句/下一句误判为本句已显示）。"""
    near = set()
    for j in range(max(0, i - span), min(len(sents), i + span + 1)):
        if j != i:
            near |= set(toks(sents[j]))
    return [t for t in toks(sents[i]) if len(t) >= 3 and t not in STOP and t not in near]


# ---------------------------------------------------------------- 音频侧对齐

def speech_times(sents, words):
    """words: [{"w":词,"s":起,"e":止}, ...]（faster-whisper 词级时间戳）
    返回每句 {"start","end","asr"} 或 None（未对齐上）。"""
    asr_tok = [norm(w["w"]).replace(" ", "") for w in words]
    ref_tok, ref_sid = [], []
    for i, s in enumerate(sents):
        for t in toks(s):
            ref_tok.append(t)
            ref_sid.append(i)
    sm = difflib.SequenceMatcher(None, asr_tok, ref_tok, autojunk=False)
    sid = [None] * len(asr_tok)
    for a, b, n in sm.get_matching_blocks():
        for k in range(n):
            sid[a + k] = ref_sid[b + k]
    last = None
    for i in range(len(sid)):
        if sid[i] is None:
            sid[i] = last
        else:
            last = sid[i]
    for i in range(1, len(sid)):
        if sid[i] is not None and sid[i - 1] is not None:
            sid[i] = max(sid[i], sid[i - 1])
    out = []
    for i, s in enumerate(sents):
        idx = [k for k, x in enumerate(sid) if x == i]
        if not idx:
            out.append(None)
            continue
        out.append({"start": words[idx[0]]["s"], "end": words[idx[-1]]["e"],
                     "asr": " ".join(words[k]["w"] for k in idx)})
    return out


def tok_times(words, sents):
    """每句每个参考词对应的 ASR 结束时间（按文本对齐；未匹配词取相邻已知词时间填补）。"""
    asr_tok = [norm(w["w"]).replace(" ", "") for w in words]
    ref, sid, pos = [], [], []
    for i, s in enumerate(sents):
        for k, t in enumerate(toks(s)):
            ref.append(t)
            sid.append(i)
            pos.append(k)
    sm = difflib.SequenceMatcher(None, asr_tok, ref, autojunk=False)
    end_t = [None] * len(ref)
    for a_, b_, n in sm.get_matching_blocks():
        for k in range(n):
            end_t[b_ + k] = words[a_ + k]["e"]
    last = 0.0
    for k in range(len(ref)):
        if end_t[k] is None:
            end_t[k] = last
        else:
            last = end_t[k]
    out = [[] for _ in sents]
    for k in range(len(ref)):
        out[sid[k]].append((ref[k], end_t[k]))
    return out


# ---------------------------------------------------------------- 显示侧（低帧率粗定位）

def load_frames(path):
    frames = []
    with open(path, encoding="utf8") as f:
        for line in f:
            line = line.strip()
            if line:
                frames.append(json.loads(line))
    return frames


def frame_text(fr, side):
    lines = sorted(fr.get(side, []), key=lambda x: x[0])
    return norm(" ".join(l[1] for l in lines))


def frame_at(frames, t):
    k = min(range(len(frames)), key=lambda i: abs(frames[i]["t"] - t))
    return frames[k]


def first_frame(frames, side, tok, t_from, t_to, base_txt):
    if tok_present(tok, base_txt):
        return None
    for fr in frames:
        if fr["t"] < t_from:
            continue
        if fr["t"] > t_to:
            break
        if tok_present(tok, frame_text(fr, side)):
            return fr["t"]
    return None


def show_times(sents, speech, frames, side, wtimes, lookahead=15.0):
    """每句该侧显示 start/end 的粗定位（低帧率采样帧序列中首次出现）。"""
    out = []
    for i, s in enumerate(sents):
        sp = speech[i]
        if sp is None:
            out.append({"start": None, "end": None})
            continue
        dist = set(distinctive(sents, i))
        seq = [(t, e) for t, e in wtimes[i] if t in dist]
        base_txt = frame_text(frame_at(frames, sp["start"]), side)
        st, st_tok = None, None
        for t, e in seq[:6]:
            ft = first_frame(frames, side, t, e - 0.3, sp["start"] + lookahead, base_txt)
            if ft is not None and (st is None or ft < st):
                st, st_tok = ft, t
        en, en_tok = None, None
        for t, e in reversed(seq[-5:]):
            ft = first_frame(frames, side, t, max(e - 0.3, st or 0), sp["end"] + lookahead, base_txt)
            if ft is not None:
                en, en_tok = ft, t
                break
        out.append({"start": st, "end": en, "head": st_tok or "", "tail": en_tok or "",
                     "dist": " ".join(t for t, _ in seq)})
    return out


# ---------------------------------------------------------------- OCR 精修（高帧率逐帧）

class OcrEngine:
    """按 (side, 秒) 缓存 OCR 结果的引擎；video/roi/y0y1 由调用方指定，不含任何硬编码路径。"""

    def __init__(self, video, roi, y0, y1, cache_path=None):
        self.video = str(video)
        self.roi = roi          # {"L": (x0,x1), "R": (x0,x1), ...}
        self.y0, self.y1 = y0, y1
        self.cache_path = Path(cache_path) if cache_path else None
        self.cache = {}
        if self.cache_path and self.cache_path.exists():
            self.cache = json.load(open(self.cache_path, encoding="utf8"))
        self._ocr = None
        self._cap = None
        self._fps = None

    def _ensure(self):
        if self._ocr is None:
            import cv2
            from rapidocr_onnxruntime import RapidOCR
            self._ocr = RapidOCR()
            self._cap = cv2.VideoCapture(self.video)
            self._fps = self._cap.get(cv2.CAP_PROP_FPS)

    def text_at(self, side, t):
        key = f"{side}:{t:.3f}"
        if key in self.cache:
            return self.cache[key]
        self._ensure()
        self._cap.set(1, round(t * self._fps))  # cv2.CAP_PROP_POS_FRAMES == 1
        ok, fr = self._cap.read()
        txt = ""
        if ok:
            x0, x1 = self.roi[side]
            res, _ = self._ocr(fr[self.y0:self.y1, x0:x1])
            lines = sorted([(b[0][1], tx) for b, tx, sc in (res or [])
                             if any(c.isascii() and c.isalpha() for c in tx)])
            txt = norm(" ".join(l[1] for l in lines))
        self.cache[key] = txt
        return txt

    def save(self):
        if self.cache_path:
            json.dump(self.cache, open(self.cache_path, "w", encoding="utf8"), ensure_ascii=False)

    def refine_time(self, side, t_found, crit, back=0.8, step=0.1):
        """t_found 为粗定位命中帧；向前按 step 再按 1/60s 找最早满足 crit 的时刻。"""
        best = t_found
        t = t_found - step
        while t >= t_found - back - 1e-6:
            if crit(self.text_at(side, round(t, 3))):
                best = round(t, 3)
                t -= step
            else:
                break
        for k in range(1, 6):
            t = round(best - k / 60, 3)
            if crit(self.text_at(side, t)):
                best = t
        return best

    def anim_start(self, side, t_text, back=0.8):
        """从文字可辨认时刻向前回溯气泡滑入动画的起点（灰度帧差，与前 2 帧比较）。"""
        import cv2
        self._ensure()
        x0, x1 = self.roi[side]
        n = int(back * self._fps) + 12
        f0 = round(t_text * self._fps)
        gs = []
        for k in range(n, -1, -1):
            self._cap.set(1, f0 - k)
            ok, fr = self._cap.read()
            if not ok:
                gs.append(None)
                continue
            g = cv2.cvtColor(fr[self.y0:self.y1, x0:x1], cv2.COLOR_BGR2GRAY)
            gs.append(cv2.resize(g, (145, 140)).astype("float32"))
        d = [float(abs(gs[i] - gs[i - 2]).mean()) if gs[i] is not None and gs[i - 2] is not None else 0.0
             for i in range(2, len(gs))]
        if not d:
            return t_text
        base = sorted(d[:12])[len(d[:12]) // 2] if d[:12] else 0.0
        thr = max(base * 3, 1.5)
        j = len(d) - 1
        miss = 0
        start_j = j
        while j >= 0:
            if d[j] > thr:
                start_j = j
                miss = 0
            else:
                miss += 1
                if miss > 2:
                    break
            j -= 1
        frame_idx = f0 - n + 2 + start_j - 2
        t = frame_idx / self._fps
        return round(t, 3) if t_text - t <= back else t_text


def refine_all(sents, speech, frames, show, side, ocr: OcrEngine):
    for i, s in enumerate(sents):
        r = show[i]
        if r.get("start") is None:
            continue
        sp = speech[i]
        base_txt = frame_text(frame_at(frames, sp["start"]), side)
        dist_head = [t for t in r["dist"].split()[:6] if not tok_present(t, base_txt)]
        tok0 = r["head"]
        if tok0 and tok0 not in dist_head:
            dist_head.insert(0, tok0)

        def crit_start(txt, _dist_head=dist_head):
            return any(tok_present(t, txt) for t in _dist_head)

        r["start_5fps"] = r["start"]
        r["start_text"] = ocr.refine_time(side, r["start"], crit_start)
        r["start"] = ocr.anim_start(side, r["start_text"])
        if r.get("end") is not None and r["tail"]:
            tok1 = r["tail"]
            r["end_5fps"] = r["end"]
            r["end"] = ocr.refine_time(side, r["end"], lambda txt, _t=tok1: tok_present(_t, txt), back=0.4)
        ocr.save()
    return show
