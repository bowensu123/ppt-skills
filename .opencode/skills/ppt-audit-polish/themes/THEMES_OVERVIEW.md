# 主题选择清单（agent 给用户看的）

当用户请求美化/重新设计 deck 但没有显式指定 `--theme`，**agent
必须先把主题选项呈现给用户，让用户选**。不要默认 auto-pick。

agent 的呈现可以是：

  1. **基于内容给出 top 3 建议**（用户接受其中一个最快）
  2. **附带"看全部 16 个"选项**（用户想看全的话）

agent 应该在询问前先读 content（或快速看一眼 deck 内容）以做出
有依据的 top-3 建议。

---

## 16 个内置主题

### 通用 5 个

| # | 名字 | 一句话风格 |
|---|---|---|
| 1 | `clean-tech` | 科技蓝（默认） — AI / SaaS / 一般科技 |
| 2 | `business-warm` | 暖橙 — 营收 / 客户 / 市场叙事 |
| 3 | `academic-soft` | 浅学术 — 普通研究 / 教学材料 |
| 4 | `editorial-dark` | 深色编辑 — 博客 / 观点 / 品牌叙事 |
| 5 | `claude-code` | 终端深色 + coral — dev / agent / CLI |

### 商业级 10 个

| # | 名字 | 一句话风格 |
|---|---|---|
| 6 | `minimalist-business` | 麦肯锡 / 贝恩 / BCG —— 大量留白、几何线条、克制黑白灰 |
| 7 | `consulting` | 专业咨询 —— 深蓝 + 酒红、信息密度高、conclusion-first |
| 8 | `modern-tech` | Stripe / Linear —— 深色 + 紫青渐变、抽象几何 |
| 9 | `corporate-classic` | 企业经典 —— 信任蓝 + 金线、传统大公司风 |
| 10 | `pitch-deck` | 投资路演 —— YC 橙 + 大字号 + 强边框、节奏感 |
| 11 | `editorial-magazine` | 杂志风 —— 暖米底 + 巨型标题、品牌发布 |
| 12 | `data-heavy` | 数据看板 —— 紧凑布局、数据系列分色 |
| 13 | `academic-research` | 学术研究 —— 衬线字体、奶油纸底、严谨 |
| 14 | `creative-agency` | 创意行业 —— 大胆撞色、强边框、异形排版 |
| 15 | `dark-premium` | 极致深色 —— 全黑 + 金银点缀、奢华大间距 |

### 品牌 1 个

| # | 名字 | 一句话风格 |
|---|---|---|
| 16 | `huawei-style` | 华为模板 —— 红色品牌色 + 微软雅黑/Arial 双字体 + **固定 title/tag/content 区域** |

---

## Agent 询问范例

### 场景 1：用户说"美化 PPT"，没指定主题

```
agent: 我看了下 deck 的内容（涉及 SaaS 产品发布 + 用户增长数据），
       推荐 3 个最契合的主题供你选：

       1) modern-tech     —— Stripe / Linear 风格，深色 + 紫青渐变
                             适合 SaaS 产品发布的视觉语言
       2) clean-tech      —— 科技蓝默认，浅色背景，最稳妥
       3) data-heavy      —— 数据看板风，如果你想突出增长指标

       直接说编号或名字就行；想看全部 16 个 → 说"全部"。
```

### 场景 2：用户说"全部"

```
agent: [显示完整的 16 项表格（上方 markdown）]

       挑一个？
```

### 场景 3：用户直接说"用 minimalist-business"

```
agent: 好，用 minimalist-business（麦肯锡风格）。
       开始美化...
       [运行 polish-business --theme themes/minimalist-business.json]
```

### 场景 4：用户已经在第一句说"用 huawei 风格"

```
agent: 直接用 huawei-style，跳过询问。
       注意：huawei-style 主题包含 layout_regions（title bar / 右上角 tag /
       content 三区固定布局），我会按规定区域排版。
       [运行 polish-business --theme themes/huawei-style.json]
```

---

## Agent 应用规则

1. **零参数调用** → 询问主题
2. **用户在初始指令里指定主题** → 直接用，不再问
3. **用户说"自动"/"随便"/"你决定"** → agent 自己看内容选最合适的（fallback 到 `pick_theme()` 关键词匹配）
4. **批处理 / CLI 模式（非 agent 驱动）** → polish-business 内部 `pick_theme()` 自动选
5. **主题里有 `layout_regions`（如 huawei-style）** → 应用 + 提醒用户该主题强约束 layout

---

## CLI 命令对应

```bash
# 显式指定主题（用户已选好）
python scripts/mutate.py polish-business --in deck.pptx --out v1.pptx \
    --theme themes/<chosen>.json

# 或 polish.py 一键流水线指定主题
python scripts/polish.py --in deck.pptx --out polished.pptx \
    --theme themes/<chosen>.json
```

---

## 主题字段速查（agent 内部参考）

每个主题 JSON 都有 5 块：

- `palette`：10 个颜色 role（primary / accent / text_strong / background / surface / border / etc.）
- `typography`：font_family + size_pt + bold + color_role 各 6 个 role
- `spacing`：5 个 EMU 间距值
- `decoration`：10 个装饰参数（card_fill_role / corner_radius / shadow / etc.）
- 可选 `layout_regions`：title / tag / content 区域（仅 huawei-style 当前有）
- 可选 `slide_dims`：强制 slide 尺寸（仅 huawei-style 显式声明）
