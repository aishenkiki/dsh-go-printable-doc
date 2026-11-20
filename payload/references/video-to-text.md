# 视频 / 音频来源：先把正文取下来

用户给的是短视频、播客、直播切片时，版面再漂亮也没用——先得有正文。

顺序固定：**有软字幕就抽字幕 → 没有就语音转写 → OCR 烧录字幕是下策**。
全文用中文写、拿不准的句子删掉，不要凭印象补。

## 1. 先判断链接是什么

```sh
curl -sIL "https://m.toutiao.com/is/XXXX/" | grep -i "^location:"
```

今日头条的 `m.toutiao.com/video/<id>/` 是视频，`/article/<id>/` 才是图文。
`https://www.toutiao.com/video/<id>/` 只有 JS 壳，正文 0 段，别在里面找文字。

## 2. 头条视频取直链

页面 HTML 里有 `<script id="RENDER_DATA" type="application/json">`，内容是 URL 编码的 JSON：

```python
import json, re, urllib.parse
html = open('/tmp/toutiao.html', encoding='utf-8').read()
raw = re.search(r'id="RENDER_DATA"[^>]*>(.*?)</script>', html, re.S).group(1)
data = json.loads(urllib.parse.unquote(raw))
```

JSON 的键是路径形状的，递归找 `videoId` / `playAuthTokenV2` 比按固定路径取省事。
`articleInfo.playAuthTokenV2` 做 `base64.urlsafe_b64decode` 后得到一段 JSON，
里面的 `GetPlayInfoToken` 就是预签名 query（`Action=GetPlayInfo&Version=2019-03-15&X-Amz-*&video_id=…&aid=13`），
直接 GET 它：

```python
import base64, json, urllib.request
token = json.loads(base64.urlsafe_b64decode(auth))['GetPlayInfoToken']
play = json.load(urllib.request.urlopen('https://vod.bytedanceapi.com/?' + token))
mp4 = play['Result']['Data']['PlayInfoList'][0]['MainPlayUrl']
```

实测一条：854×480、h264+aac、88 秒、8.3 MB，无字幕轨。

## 3. 有软字幕就先抽字幕（最省事）

```sh
ffprobe -v error -show_entries stream=codec_type,codec_name -of csv=p=0 in.mp4
# 0,h264,video
# 1,aac,audio          ← 只有这两条 = 没有字幕流，字幕是烧在画面上的
ffmpeg -y -i in.mp4 -map 0:s:0 out.srt      # 真有字幕流时才跑这句
```

## 4. 没有字幕 → 语音转写（推荐路线）

抽音轨：16k 单声道就够，别留立体声大文件。

```sh
ffmpeg -y -i in.mp4 -vn -ac 1 -ar 16000 -f wav audio.wav
pip install sherpa-onnx
```

模型用 SenseVoice（中/英/日/韩/粤，int8 版 229 MB）：

```sh
# GitHub 直连会被 reset；hf-mirror 实测 3 分钟只下 16K，别用
curl -L -C - -o sv.tar.bz2 \
  "https://gh-proxy.com/https://github.com/k2-fsa/sherpa-onnx/releases/download/asr-models/sherpa-onnx-sense-voice-zh-en-ja-ko-yue-2024-07-17.tar.bz2"
apt-get install -y bzip2        # 少了它 tar 会报 "lbzip2: Cannot exec"
tar xf sv.tar.bz2
```

识别（`OfflineStream` **没有** `accept_wave_file`，只能喂数组）：

```python
import wave, numpy as np, sherpa_onnx

w = wave.open('audio.wav')
sr = w.getframerate()
pcm = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16).astype(np.float32) / 32768

rec = sherpa_onnx.OfflineRecognizer.from_sense_voice(
    model='…/model.int8.onnx', tokens='…/tokens.txt',
    num_threads=4, use_itn=True, language='zh')
s = rec.create_stream()
s.accept_waveform(sr, pcm)
rec.decode_stream(s)
print(s.result.text)
```

实测：88 秒音频整段一次识别 **4.6 秒**，输出是带标点的连续文本，能直接整理成稿。
`use_itn=True` 会加标点、把数字写成阿拉伯数字，省掉后处理。

- 有一句没听清就按转写走，不要按印象补；
- 同音错字（人名、成语）成稿时顺手改，但别改内容；
- 实在拼不出的句子整句删，不要编。

## 5. OCR 烧录字幕为什么是下策

本机实测（tesseract 5.3.4 + chi_sim，psm 6）：

| 项 | 数字 |
| --- | --- |
| 单帧识别 | 0.50 秒 |
| 2fps × 88 秒 = 176 帧 | 约 89 秒，还不含掩膜、去重、拼句 |
| 准确度 | 大量乱码：「老师第一个叫的就是你」→ `一老师第ae介就是叫你[ET下` |

中文 + 卡通背景 + 白字黑边 + 854×480 + 一句话被拆在几帧里，五个因素叠一起就没法成稿。
灰阶、二值化、3 倍放大、只留白色像素都试过，只是乱码换了个样子。

只有「高清、标准字体、每句有完整停留帧」的视频，OCR 才可能比转写省事。

## 6. 成稿时写清来源

页脚写平台 + 账号 + 标题，例：

```
整理自 今日头条「知然见闻」·《成绩好到底有多爽？这5种快乐，只有经历过的人才懂》
```

**不要拿同题的其他文章冒充原文**——标题一样、内容泛泛，读者一对照就露馅。
转写整理出的稿子，措辞同义改写可以，事实与观点必须来自这条音视频。

## 7. 环境（本机踩过的）

- `apt-get install ffmpeg` 前先 `apt-get update`：列表过期会在 `alsa-ucm-conf` 上报 404。
- 模型下载断了用 `curl -C -` 续传；本机分了 3 轮（314M → 911M → 999M → 1047870769 字节）。
- `/tmp` 会被环境重置清掉，模型要重下。长期用就挪到 `~/.cache/` 之下再指过去。
