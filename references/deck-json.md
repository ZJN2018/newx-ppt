# deck.json 格式（整稿）

一份 `deck.json` 就是一整套 PPT：整稿设置、目录章节、正文页、结束页。正文页的写法见 [版面切分规格](split.md) 第六节，组件参数见 [组件目录](components.md)。

```json
{
  "settings": {
    "title": "2026 年度生产运营工作总结",
    "subtitle": "交付稳住了，成本和人才是明年的两道坎",
    "org": "生产运营部",
    "date": "2026 年 12 月",
    "preset": "商务蓝",
    "corner": "round",
    "footer": "资料来源：部门年度台账",
    "photos": {"cover": "工业、厂房、factory production line"},
    "filename": "生产运营部_年度总结"
  },
  "sections": [
    {"name": "年度概况", "desc": "交付、质量、成本三项主指标", "photo": "仓储"},
    {"name": "重点工作", "desc": "五项专项的进展与成效"}
  ],
  "pages": [ { "section": "年度概况", "title": "……", "separation": "纯卡片", "partition": { } } ],
  "closing": {"words": "感谢聆听"}
}
```

## settings

| 字段 | 必填 | 说明 |
|---|---|---|
| `title` / `subtitle` | 是 / 否 | 封面主标题，**要带单位名**（如「青岚储能2025年度工作汇报」，部门可放第二行）；副标题写一句总判断。主标题一行放不下时自动在标点、年份等数字、"与/和/及"、单位名之后断成两行，不在词中间断；要指定断行位置可在标题里写 `\n` |
| `org` / `date` | 否 | 封面单位与日期 |
| `reporter` | 否 | 汇报人（姓名或"部门＋姓名"），写在结束页"感谢聆听"下方；材料没给就不写，结束页改列 `org`，不要编造 |
| `preset` | 否 | 色系名（见 [配色](palette.md)），或 LOGO 原色值如 `"#E30613"`。照抄 `confirm.py` 记录的结果，不自己改 |
| `logo_color` | 否 | 没写 `preset` 时直接用这个 LOGO 主色推导；过灰、过深时退回最接近的预设 |
| `industry` | 否 | 前两项都没有时，按行业在色库候选里按标题挑一种；都没有用藏蓝 |
| `corner` | 否 | `square`（默认）或 `round`，整稿统一 |
| `footer` | 否 | 每页底部的来源与口径说明（左对齐一行，页码单独居中在下面）；单页可用页面的 `footer` 覆盖。格内的口径、统计范围、未经审计等说明都写到这里 |
| `styles` | 否 | **一般不写**：首页、目录、转场版式由脚本按封面标题随机抽一种，套内统一，重跑不变。只有用户明确要求某种版式时才写，如 `{"cover": "fade", "agenda": "photo_left", "transition": "full"}`，见 [结构页版式](structure-pages.md) |
| `styles_from` | 否 | 写了 `styles` 时必须同时写 `"user"`，表示是用户指定的；不写则 `styles` 在构建时被忽略 |
| `logo` | 否 | 单位 LOGO 的本机图片路径（PNG 透明底最好，JPG 也可）。有就放：封面、目录、转场、结束页右上角，正文页眉右上角小号。白底、黑底自动抠成透明；按落点背景挑对比度够的一版：浅色背景放原色（看不清时换深色剪影），深色背景放纯白反白版；反白会丢信息（实心色块、以白色笔画为主）时深色背景放原图；任何情况都不垫白色底卡；不改色相。用户给了或能取到公开 LOGO 就写，没有就不写 |
| `photos.agenda` | 否 | 目录页照片（目录用照片版式时）；不写用封面照片 |
| `photos.cover` | 否 | 封面配图：图库文件名、本机图片路径或中文关键词（顿号隔开）；不写或找不到相关的图则封面不放图 |
| `photos.online` | 否 | 默认 `true`：自带图库没有强相关的图就到开源图库检索（Openverse、Wikimedia Commons，都不需要密钥），来源与授权记入 `图片来源.json`；写 `false` 只用自带图库 |
| `photos.sources` | 否 | 联网检索的图库及顺序，如 `["wikimedia", "openverse"]`；不写则依次尝试 |
| `filename` | 否 | 输出文件名（不含扩展名），默认用 `title` |
| `public_output` | 否 | `true` 时对外交付清理：字体统一为 Noto Sans CJK SC，去缩略图 |

## sections

目录里的章节，3–6 个。`name` 与正文页的 `section` 对应，正文页进入新章节时自动插入转场页。`desc` 是一句陈述式说明；`photo` 是转场页配图（写法同 `photos.cover`），可省；找不到相关的图则不放。

## pages

每项是一页正文，字段见 [版面切分规格](split.md) 第六节。常用：`section`、`title`、`separation`、`partition`（或创意型的 `creative`）、`footer`。

格子上除 `role` / `fill` / `title` / `evidence` / `element` 外，可用 `card` 指定卡片样式（`plain` / `header` / `dark` / `number` / `label`），见 [组件目录](components.md)。

## closing

结束页与封面用同一张图、同样的主色蒙版。大字默认「感谢聆听」，下方一行默认「汇报人：`settings.reporter`」，没有汇报人时列 `settings.org`。一般不用写 `closing`；`words` / `sub` 只在用户另有要求时覆盖。

## 生成

```sh
python3 scripts/check_spec.py deck.json --evidence evidence.md   # 只检查
python3 scripts/build.py deck.json --out out/ --evidence evidence.md
```

`build.py` 先跑规格检查，有错误就停；通过后生成 `<filename>.pptx`、PDF、逐页 `-p-N.png` 和 `-overview.png`，并在台账 `ppt-ledger.jsonl` 追加一条记录。
