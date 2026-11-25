"""图标：随 skill 分发的 Lucide 线性图标（ISC 许可，见 assets/icons/LICENSE-lucide.txt），按主题色渲染成 PNG。

用法：S.icon("份额")、S.icon("trophy") 返回已按主色上色的 PNG 路径；按中文关键词挑选时取第一个命中的图标。
创意组件的图片槽位可直接传 {"icon": "关键词", "meaning": "..."}，放置时自动换成主色图标。"""
import hashlib, os, re, subprocess, tempfile
from pathlib import Path

DIR = Path(__file__).resolve().parents[1] / "assets" / "icons"
CACHE = Path(tempfile.gettempdir()) / "niuppt-icons"

KEYWORDS = {
    "trophy": "份额 第一 领先 冠军 排名", "award": "荣誉 奖项 品牌 认证", "medal": "第一 金牌 头部", "crown": "龙头 头部 领先",
    "target": "目标 定位 聚焦 战略", "flag": "里程碑 计划 重点 方向", "rocket": "增长 启动 突破 扩张", "star": "亮点 核心 重点",
    "trending-up": "增长 上升 提升 改善", "trending-down": "下降 回落 减少 承压", "chart-column": "收入 规模 业绩 柱",
    "chart-line": "趋势 走势 变化", "chart-pie": "结构 占比 构成", "percent": "比率 毛利率 费用率 比例", "gauge": "效率 指标 监测 利用率",
    "wallet": "现金 资金 现金流", "banknote": "收入 回款 货款", "coins": "利润 收益 分红", "piggy-bank": "存款 存单 储备",
    "landmark": "银行 金融 借款 融资", "receipt": "费用 税 所得税 账单", "credit-card": "支付 结算", "hand-coins": "股东回报 分红 回报",
    "circle-dollar-sign": "金额 价值 估值", "briefcase": "业务 经营", "building-2": "公司 总部 机构", "factory": "产能 工厂 生产 制造",
    "warehouse": "库存 存货 仓储", "store": "渠道 门店 零售 商超 经销", "shopping-cart": "消费 购买 电商 商超", "shopping-bag": "零售 零食店 会员店",
    "package": "产品 新品 包装", "boxes": "库存 产品线 品类", "truck": "物流 配送 供应链", "ship": "出口 海外 贸易", "container": "进出口 关税 贸易",
    "globe": "国际化 全球 海外", "earth": "全球 国家 地区", "map-pin": "区域 地区 网点 布局", "handshake": "合作 客户 伙伴 经销商",
    "users": "会员 用户 客户 团队", "user-check": "会员 私域 用户", "sparkles": "新品 创新 亮点", "lightbulb": "创新 研发 想法",
    "flask-conical": "研发 实验 配方", "microscope": "研究 研发 检测", "milk": "液体乳 牛奶 乳品 奶源", "ice-cream-cone": "冷饮 雪糕 冰淇淋",
    "baby": "婴配粉 婴幼儿 奶粉", "heart-pulse": "健康 营养 保健", "pill": "营养品 保健品 功能", "leaf": "低碳 绿色 有机 可持续",
    "sprout": "牧场 种植 新业务 培育", "wheat": "原料 农业 奶源", "chef-hat": "餐饮 厨房 烘焙", "utensils": "餐饮 食品 餐厅",
    "cpu": "数智化 数字化 科技 AI", "monitor-smartphone": "数字化 线上 系统", "smartphone": "线上 移动 即时零售", "database": "数据 数据库",
    "bot": "AI 人工智能 智能", "shield-check": "合规 安全 质量 稳健", "shield-alert": "风险 预警 减值", "triangle-alert": "风险 关注 警示",
    "scale": "口径 平衡 对比", "file-search": "审计 核查 披露", "search": "关注 检视 排查", "clipboard-check": "检验 验收 监测",
    "list-checks": "清单 任务 事项", "calendar": "时间 日程 季度", "clock": "周期 时效 账期", "layers": "层级 结构 组合",
    "network": "网络 体系 渠道网络", "settings": "运营 管理 机制", "zap": "效率 提速 能源", "refresh-cw": "周转 循环 更新",
    "recycle": "循环 环保 回收", "link": "协同 联动 整合", "sun": "季节 旺季 光伏 太阳能", "snowflake": "冷链 低温",
    "car": "汽车 整车 乘用车 新能源车 销量", "bus": "商用车 客车 公交", "battery-charging": "电池 动力电池 储能 充电",
    "battery-full": "电池 容量 续航", "plug-zap": "充电 闪充 补能 充电桩 站点", "fuel": "燃油 能源 加注",
    "train-front": "轨道交通 云轨", "cable": "线束 连接 电网", "anchor": "港口 航运 海运", "package-check": "出货 交付 发货",
    "chart-no-axes-combined": "综合 经营 指标", "calculator": "测算 核算 财务", "file-text": "报告 公告 文件",
    "gavel": "法规 合规 诉讼", "building": "物业 资产 楼宇", "map": "地图 版图 市场布局", "trending-up-down": "波动 分化 起伏",
}

def names():
    return sorted(p.stem for p in DIR.glob("*.svg"))

def pick(query):
    """按英文名或中文关键词挑图标；找不到时报错，避免悄悄用错图。"""
    q = query.strip()
    if (DIR / f"{q}.svg").exists():
        return q
    best = [(sum(1 for w in KEYWORDS.get(n, "").split() if w and (w in q or q in w)), n) for n in names()]
    score, name = max(best)
    if score == 0:
        words = sorted({w for n in names() for w in KEYWORDS.get(n, "").split()})
        raise ValueError(f"没有与“{query}”匹配的图标。可用中文关键词：{'、'.join(words)}；也可直接写英文名：{', '.join(names())}")
    return name

def _svg_to_png(src, out, size):
    """SVG 转 PNG：依次尝试 rsvg-convert、cairosvg、PyMuPDF，都没有时说明怎么装。"""
    import shutil
    if shutil.which("rsvg-convert"):
        subprocess.run(["rsvg-convert", "-w", str(size), "-h", str(size), "-o", str(out), str(src)], check=True)
        return
    try:
        import cairosvg
        cairosvg.svg2png(url=str(src), write_to=str(out), output_width=size, output_height=size)
        return
    except ImportError:
        pass
    try:
        import fitz
    except ImportError:
        raise RuntimeError("图标渲染需要 librsvg（brew install librsvg）、cairosvg 或 PyMuPDF（pip install pymupdf）之一")
    with fitz.open(str(src)) as doc:
        page = doc[0]
        page.get_pixmap(matrix=fitz.Matrix(size / page.rect.width, size / page.rect.height), alpha=True).save(str(out))

def render(query, color, size=256, stroke=2.0):
    """把图标渲染成指定颜色的 PNG，返回路径。"""
    name = pick(query)
    color = color.lstrip("#")
    CACHE.mkdir(parents=True, exist_ok=True)
    out = CACHE / f"{name}-{color}-{size}-{stroke}.png"
    if not out.exists():
        svg = (DIR / f"{name}.svg").read_text()
        svg = svg.replace("currentColor", f"#{color}")
        svg = re.sub(r'stroke-width="[\d.]+"', f'stroke-width="{stroke}"', svg)
        src = CACHE / f"{name}-{color}.svg"; src.write_text(svg)
        _svg_to_png(src, out, size)
    return str(out)
