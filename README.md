# xiaozhi-local-music-mcp

面向小智 AI 的本地音乐播放器 MCP Server，基于 Windows Media Player，支持播放控制、播放列表、随机播放和系统音量控制。

## 功能

* 本地音乐播放：支持 MP3、FLAC、WAV、M4A、AAC、OGG、WMA、OPUS
* 播放控制：播放、暂停、恢复、停止、上一首、下一首
* 普通播放自动连播
* 随机播放指定数量歌曲，支持自动连播和手动切换
* 本地音乐库扫描与歌曲列表
* 查询当前歌曲、播放进度和播放状态
* 调节 Windows 系统音量

## MCP Tools

| Tool                        | 功能     | 示例         |
| --------------------------- | ------ | ---------- |
| `play_song`                 | 播放歌曲   | "播放《一路生花》"   |
| `play_folder`               | 播放文件夹  | "播放我的音乐文件夹"  |
| `random_play`               | 随机播放   | "随机播放 10 首歌" |
| `stop_random_play`          | 停止随机播放 | "停止随机播放"     |
| `play_random_song`          | 随机播放一首 | "随机放一首歌"     |
| `pause_music`               | 暂停     | "暂停音乐"       |
| `resume_music`              | 恢复     | "继续播放"       |
| `stop_music`                | 停止     | "停止音乐"       |
| `play_next_song`            | 下一首    | "播放下一首"      |
| `play_previous_song`        | 上一首    | "播放上一首"      |
| `get_current_music_info`    | 获取播放信息 | "现在播放的是什么歌？" |
| `list_local_songs`          | 列出歌曲   | "有哪些歌曲？"     |
| `scan_local_music_folder`   | 扫描音乐目录 | "扫描我的音乐"     |
| `set_windows_system_volume` | 设置系统音量 | "把音量调到 50"  |
为避免与自带音乐mcp冲突，可以强调本地电脑

## 项目结构

```text
xiaozhi-local-music-mcp/
├── mcp_pipe.py
├── music_player.py
├── mcp_config.json
├── README.md
└── requirements.txt
```

## 环境要求

* Python 3.11+

安装依赖：

```powershell
pip install -r requirements.txt
```

默认音乐目录：

修改`DEFAULT_MUSIC_DIRECTORY`为你本地电脑的音乐目录
```text
DEFAULT_MUSIC_DIRECTORY = r"<YOUR_MUSIC_DIRECTORY>"
```

## 运行

```powershell
python mcp_pipe.py
```
