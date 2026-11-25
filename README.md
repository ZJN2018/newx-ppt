# NEWX PPT

为打工人做工作汇报的 PPT 技能（agent skill）。输入报告、数据或要点，输出可编辑的原生 PPTX，附 PDF 与逐页预览图。

- **切分树排版**：每页正文区按材料关系切成 2–6 格，每格放一个组件；四种区隔方式（纯卡片、纯线条、线卡混用、创意型）在整稿内交替。
- **配色直接指定**：有 LOGO 直接用 LOGO 原色，浅底、灰阶、渐变、强调色按规则推导（规则取自 42 套实测参考配色）；没有 LOGO 按行业从 28 种主色里挑一种，同一行业的不同稿子颜色不同。不让用户输入色值。
- **卡片与卡内小逻辑**：五种卡片样式（浅底、表头条、深色重点、编号、行标签），以及标签胶囊、提示条、图标步骤、大数字对、卡内分区等组件，文字多的页也有图示。
- **图标与图片**：自带 IconPark 线性图标（Apache-2.0，转为可改色的矢量形状）与 50 张 CC0 照片（按人工看图写的中文匹配词选图）；可选联网从 Openverse、Wikimedia Commons 取图（不需要密钥），来源与授权逐张记录。页内图片自动降饱和并向主色调色；正文图片至少挨着正文区一角。
- **两次确认**：先确认一条故事线，再询问是否补充材料；论据写进底稿，页面上的数字只能来自底稿。
- **几何交给代码**：模型写 `deck.json`（内容、切分、组件），脚本做规格检查、绘制、渲染成图；看图自查最多三轮。

## 用法

把整个目录放进 agent 的 skills 目录，agent 读 `SKILL.md` 后按流程执行。也可以直接从一份 `deck.json` 出稿：

```sh
python3 scripts/build.py examples/生产运营部_文字多.deck.json --out out/
```

`deck.json` 的格式见 [deck.json 格式](references/deck-json.md) 与 [版面切分规格](references/split.md)；`examples/` 里有数据多、文字多和组件样张三份示例。

## 类型

| 类型 | 状态 |
|---|---|
| 工作汇报 | 可用 |
| 其他类型 | 计划中 |

## 安装

将 ZIP 解压，把完整的 `newx-ppt/` 文件夹放进 agent 的技能目录。

**从旧版升级**：技能目录里如果已有 `niuma-ppt/`（小牛 PPT、牛马 PPT 时期的旧版），先把它移出技能目录或删除，再放入 `newx-ppt/`。新旧两版的触发描述相同，同时留在技能目录里，制作 PPT 时可能调用到旧版。旧稿的 deck.json 可以直接用新版重新构建。

下面的步骤在**运行技能的电脑**上执行。只阅读交付的 PDF 可以直接使用 PDF 阅读器。接收方要在 PowerPoint 编辑 PPTX，建议也安装稿件使用的字体，减少替换字体造成的换行变化。

### 先装 LibreOffice 和中文字体

- 安装 Python 3.10 或以上和 LibreOffice。技能用 LibreOffice 将 PPTX 转成 PDF 再逐页看图。macOS 按电脑选择 Apple Silicon 或 Intel，将 LibreOffice 拖入“应用程序”。Windows 使用对应架构的安装包，按默认路径安装即可。[Python 下载](https://www.python.org/downloads/) · [LibreOffice 下载](https://www.libreoffice.org/download/)
- 默认稿件的正文和数字用 **思源黑体 CN**，标题用 **思源宋体 CN**。下载简体中文 CN 静态 OTF，优先安装两套字体的 Regular、Bold。macOS 双击字体后在“字体册”点安装，Windows 解压后选中 OTF、右键“安装”。[思源黑体 CN](https://github.com/adobe-fonts/source-han-sans/tree/release/SubsetOTF/CN) · [思源宋体 CN](https://github.com/adobe-fonts/source-han-serif/tree/release/SubsetOTF/CN)
- 在 deck.json 的 settings 写 `"public_output": true` 对外交付时，PPTX 内字体统一为 **Noto Sans CJK SC**，请在制作机安装该字体。缺少指定字体时系统可能替换字体，窄栏换行会有变化。[Noto 简体中文黑体](https://github.com/notofonts/noto-cjk/tree/main/Sans/OTF/SimplifiedChinese)

### macOS 终端

在解压后的技能目录执行。第三行安装逐页出图需要的 PyMuPDF。如果系统已有 Poppler 的 `pdftoppm`，可以跳过第三行。

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip install pymupdf
.venv/bin/python scripts/check_install.py
```

### Windows PowerShell

在解压后的技能目录执行。直接使用虚拟环境中的 Python 即可。若 Python 安装器提供的命令是 `python`，第一行可使用 `python -m venv .venv`。已有 `pdftoppm` 时同样可跳过第三行。

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install pymupdf
.\.venv\Scripts\python.exe scripts\check_install.py
```

看到“安装检查通过”后，让 agent 使用上述虚拟环境的 Python。要看实际字体，在检查命令末尾加 `--out install-preview`（使用新的目录名），打开生成的中文样页对照标题和正文。检查显示的字体路径是绘制预览标签选用的字体，PPTX 的正文字体和标题字体按上述两种输出模式确定。

若检查提示找不到 LibreOffice，先确认应用已安装。自定义位置可在当前终端设置 `NIUPPT_SOFFICE` 为可执行文件的完整路径，再运行同一检查。

```sh
# macOS
export NIUPPT_SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
```

```powershell
# Windows PowerShell，自定义安装时换为实际位置
$env:NIUPPT_SOFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'
```

## 素材与版权

NEWX PPT 的代码与文档以 [MIT 许可证](LICENSE) 发布。

自带图标来自 [IconPark](https://github.com/bytedance/IconPark)（Apache-2.0，许可证见 `assets/LICENSE-IconPark.txt`）；自带照片为 CC0。联网取图来自多个开源图库，授权逐张不同，来源与授权记入 `图片来源.json`，使用前请逐张核查。用户自行替换的图片请确认有使用权，详见 [素材与版权](references/assets.md)。
