# latency_toolkit — 字幕/转录显示时延自动测量工具

测两台设备（比如开发环境手机 vs 测试环境手机）的字幕/转录显示，相对原始语音落后多少秒。
全自动：语音识别定位每句话说完的时间，OCR 扫屏幕定位每句字幕出现的时间，两者相减就是时延，结果导出成 Excel。

## 1. 安装

需要 Python 3.9+。

```bash
pip install -r requirements.txt
```

首次运行会自动下载 faster-whisper 和 RapidOCR 的模型文件（几十到几百 MB，需联网）。

## 2. 准备两个东西

- **视频**：录屏/录像文件，里面同时拍到两台设备的屏幕（或一台设备但要对比的两个区域）。
- **台词文本** `sentences.txt`：每行一句原始台词（跟视频里说的话一致，用于对齐语音识别结果）。

## 3. 一键跑全流程

```bash
python cli.py all \
  --video "D:/my_test.mp4" \
  --sentences sentences.txt \
  --roi L=300,880 --roi R=920,1520 \
  --y0 400 --y1 960 \
  --out result.xlsx
```

- `--roi 名字=x0,x1`：屏幕在视频画面里的左右像素范围，可以传多个（不止两侧也行，比如 `--roi A=... --roi B=... --roi C=...`）。
- `--y0/--y1`：字幕所在的上下像素范围。
- 先用任意看图工具打开视频某一帧截图，量出各设备字幕区域的像素坐标即可。

跑完会生成：
- `result.xlsx`：每句的语音起止、各侧显示起止、时延，以及平均值
- `result.words.json` / `result.frames.jsonl`：中间结果，重复跑同一个视频时可以复用（不传 `--force` 就自动跳过重算）

## 4. 分步跑（进阶，比如换个 sentences.txt 重新算不想重跑识别）

```bash
python cli.py asr --video my.mp4 --out words.json          # 语音识别
python cli.py ocrscan --video my.mp4 --roi L=300,880 --roi R=920,1520 --y0 400 --y1 960 --out frames.jsonl   # OCR 扫描
python cli.py calc --sentences sentences.txt --words words.json --frames frames.jsonl --out result.xlsx      # 计算时延
```

`calc` 默认按视频里低帧率（默认 5fps，由 `ocrscan --step` 控制）采样的结果算，误差在一两个采样间隔内。
如果要精确到 1/60 秒，加 `--refine --video my.mp4 --roi L=300,880 --roi R=920,1520 --y0 400 --y1 960`，
会对命中附近重新逐帧解码 OCR，速度慢很多（几分钟到几十分钟，视频越长越慢），但结果会精确到帧。

## 5. 如果你还用剪映做人工标记对比

`marks` 子命令单独读取剪映草稿里的标记点（复合片段内的 time_marks），不依赖上面的语音/OCR 流程：

```bash
python cli.py marks --draft "草稿名关键词" --out 剪映标记时间差.xlsx
```

前置操作（剪映内，仅一次）：选中带标记的片段 → 右键「新建复合片段」→ Ctrl+S 保存；之后每次改完标记再 Ctrl+S 即可重新读取。

标记颜色约定（`marks.py` 顶部 `COLOR_KEY`/`CYAN` 可自行改成你自己的配色）：

| 颜色 | 含义 |
|---|---|
| 青色 `S{n}` | 第 n 句音频句首 |
| 红色 `S{n}L` | 环境 A（左）显示句首 |
| 橙色 `S{n}R` | 环境 B（右）显示句首 |
| 青色 `T{n}` | 第 n 句音频句尾 |
| 黄绿色 | 环境 A 显示句尾 |
| 紫色 | 环境 B 显示句尾 |

`calc` 加上 `--draft 草稿名关键词` 可以在同一份 Excel 里带出对应的人工标记编号，方便交叉核对。

## 6. 常见问题

- **识别慢/不准**：`asr --model` 默认 `small`（CPU 上较快），换成 `medium` 或 `large-v3` 更准但更慢；有 GPU 可加 `--device cuda`。
- **OCR 认不出字**：先确认 `--roi`/`--y0`/`--y1` 截取的区域刚好框住字幕文字，太大会把无关内容也 OCR 进去。
- **两台设备不是左右分布**：`--roi` 名字和坐标随意起，几侧都支持，不局限于 L/R。
