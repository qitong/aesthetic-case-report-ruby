# Custom Design Specification (design.md)
# 医美 Case Report PPT 自定义设计规范范例

当您希望自定义幻灯片的视觉主题、色彩体系、字体或指定页面结构时，可提供或修改本 `design.md` 文件。未在文件中指定的属性将自动沿用 `aesthetic-case-report-ruby` 的高审美默认模板规范。

---

## 1. 演示文稿元数据 (Deck Metadata)
- **Title (主标题)**: 面部全层抗衰与眶周精细化两步融合技术 · 临床个案精研
- **Subtitle (副标题)**: 结合解剖间隙机制与纵向 Patient Journey 随访观察
- **Speaker (讲者)**: 刘志昱 医生 (Dr. Zhiyu Liu)
- **Institution (机构/会议)**: 美学临床实践学术研讨会
- **Target Slides (预期页数)**: 14–18 页
- **Aspect Ratio (画幅比例)**: 16:9 (13.333 × 7.500 inches)

---

## 2. 视觉与色彩令牌 (Theme & Color Tokens)

支持指定深浅双模态画布、主题色、强调色及中英文字体（构建引擎会自动解析以下 YAML 代码块并应用至整套 PPTX）：

```yaml
theme:
  style_name: "Ruby Clinical Gallery"
  bg_dark: "#111216"          # 深色画布背景（用于封面、临床照片页、Patient Journey 走廊）
  bg_card_dark: "#1C1D24"     # 深色页卡片容器底色
  border_dark: "#2E303D"      # 深色页卡片描边色
  bg_light: "#F5F5F7"         # 浅色画布背景（用于文献证据页、多维卡片页、评估档案页）
  bg_card_light: "#FFFFFF"    # 浅色页卡片容器底色
  border_light: "#D2D2D7"     # 浅色页卡片与表格边框色
  text_on_dark: "#F5F5F7"     # 深色背景主文字色
  muted_on_dark: "#A1A1A6"    # 深色背景次要文字与图注色
  text_on_light: "#1D1D1F"    # 浅色背景主文字色
  muted_on_light: "#6E6E73"   # 浅色背景次要文字与页脚色
  accent_primary: "#C85A70"   # 主强调色（宝石红/珊瑚红：用于眉标、解剖重点1、核心标题装饰）
  accent_secondary: "#C5A880" # 辅强调色（香槟暖金：用于副标题、时间节点、二期微调标签）
  accent_blue: "#4EA3FF"      # 语义蓝（用于韧带结构 ORL/TTL、第一步支撑标识）
  accent_emerald: "#36D47E"   # 语义绿（用于安全间隙沉积、术后改善高亮胶囊框）
  font_cn: "PingFang SC"      # 中文字体（如 PingFang SC, Source Han Serif CN, Microsoft YaHei）
  font_en: "Helvetica Neue"   # 英文字体（如 Helvetica Neue, Arial, Constantia）
```

> **其他常用预设配色灵感**：
> - **Galderma Executive Navy（高德美深海蓝）**：`bg_dark: "#0B192C"`, `accent_primary: "#00A3C4"`, `accent_secondary: "#C5A880"`
> - **Academic Monochrome（极简黑白灰学术风，参考 Blending 精读）**：`bg_dark: "#111113"`, `accent_primary: "#FF5B4F"`, `accent_secondary: "#4EA3FF"`
> - **Aesthetic Plum & Rose（典雅紫金抗衰风，参考 Aesthetic Lives）**：`bg_dark: "#121116"`, `accent_primary: "#A8A4D9"`, `accent_secondary: "#C18668"`

---

## 3. 页面结构与版式编排要求 (Slide Structure & Layouts)

可用标准版式标识（`layout`）：
- `cover`：全屏深色学术封面
- `section`：大号数字章节过渡页
- `cards`：2–3 列结构化卡片页（适用于临床痛点、全层抗衰机理、文献核心论断）
- `anatomy` / `narrative`：左图右文解剖机制与临床问题精析页（支持红/蓝/绿/金语义高亮）
- `profile`：患者基线档案与术前量化评估（双栏对称卡片）
- `gallery`：术前多角度基线影像展陈（1–3 张同批照片等比呈现）
- `treatment`：面部画线布点图 + 原生可编辑精细剂量表（部位 / 侧别 / 层次 / 工具 / 产品 / 剂量）
- `comparison`：2-Up 同角度前后对比 + 右侧临床观察要点卡片
- `matrix`：4-Up 静态与微笑动态（或四视角）对比矩阵
- `journey`：3–5 时点横向 Patient Journey 时光走廊
- `evidence`：双栏文献证据评价与本例对照（左栏文献事实/结局，右栏本例对应与证据局限）
- `closing`：临床要点总结、讨论问题与合规致谢页

### 自定义页序示例：
1. **[cover]** 封面：主副标题、讲者与会议保密声明
2. **[narrative]** 临床痛点：泪沟与中面部填充后的动态失真问题
3. **[anatomy]** 解剖机制：弓状缘（AM）、泪沟韧带（TTL/ORL）与眶隔前间隙
4. **[cards]** 分型评估与两步注射策略（先深层支撑 DMCFP，后浅层穿韧带融合）
5. **[profile]** 患者基线档案与术前 FAS 面部评估
6. **[gallery]** 术前正位、45° 斜位、90° 侧位基线照片
7. **[treatment]** 首次治疗布点图与左右侧精确注射参数表
8. **[comparison]** 首次治疗前后对比（正位静态 + 观察要点）
9. **[comparison]** 首次治疗前后对比（45° 斜位 / 90° 侧位）
10. **[treatment]** 6 个月随访再评估与二次治疗微调方案（如有）
11. **[journey]** Patient Journey 全程多时点横向对比走廊
12. **[evidence]** 文献证据边界与本例临床实践反思
13. **[closing]** 临床核心收获与合规免责声明

---

## 4. 视觉与医学合规红线 (Visual & Clinical Guardrails)
- **照片零变形**：所有患者照片必须保持原始宽高比（等比缩放居中），严禁拉伸或镜像翻转。
- **深底衬托临床照**：凡包含患者面部对比照片的页面，统一采用 `bg_dark` 深色背景，消除白底眩光对肤色和凹陷阴影判断的干扰。
- **底部合规声明**：封面与内页底部保留学术交流与产品说明书合规声明。
