---
name: aesthetic-case-report-ruby
description: 将医美领域的患者病程记录（patient journey）、实际治疗情况（靶区、层次、针具、产品、左右侧剂量）和治疗前后对比照片制作成高美学标准的 16:9 可编辑 Case Report PPTX，支持有或无参考学术论文（PDF/文献），支持选用总结自经典案例的内置 Default 模板或用户提供的自定义 design.md。当用户提到“医美 case report”、“病例汇报 PPT”、“patient journey PPT”、“结合论文做病例 PPT”、“前后对比照做 PPT”或调用 aesthetic-case-report-ruby 时使用。
---

# Aesthetic Case Report Ruby (医美临床病例与学术机制 PPT 生成器)

本技能专用于生成医美领域具有国际学术会议与高级临床美学水准的可编辑 **16:9 PPTX 病例汇报**。技能提炼并融合了两套经典范式：
1. **文献机制结合临床案例型 (`evidence-case`)**：提炼自《The Blending Technique: A New Approach for Tear Trough Correction》论文及配套精读 PPT，将“临床痛点 → 分层解剖与间隙机制 → 分型决策树 → 两步/分层操作与流变学选品 → 真实病例验证 → 文献证据边界与批判”融为一体。
2. **纵向全层抗衰旅程型 (`journey`)**：提炼自《Aesthetic Lives (Dr Jovi Liu Edition)》，以真实患者跨周期就诊旅程为主线，聚焦“全层抗衰评估（FAS 量表） → 首次治疗面部布点与左右侧精确剂量 → 多角度前后对比（正/斜/侧位/动态） → 阶段随访与二期微调 → 多时点 Patient Journey 横向时光走廊”。

---

## 1. 盘点输入与结构化问询 (Input Discovery & Intake)

当用户发起任务时，先读取用户已提供的病程描述、文件或路径，**仅对缺失项发起一次精炼问询**（若当前环境支持 `ask_question` 工具，优先使用结构化多选/单选交互；否则使用清晰的列表询问），明确以下 4 项核心配置：

1. **输入素材类型盘点**：
   - **必选三要素**：
     - ① **患者病程与就诊旅程 (Patient Journey)**：代号/年龄、主诉、既往医美史、术前评估/分型、随访反馈。
     - ② **实际治疗情况 (Treatment Records)**：治疗次数與时间点、解剖靶区、患者左/右侧别（Left / Right / Midline）、注射层次、针具规格、产品/设备及**实际剂量（mL / U / 瓶数）**。
     - ③ **前后对比照片所在文件夹 (Photos Folder)**：包含术前基线、术前画线布点图、术后即刻、各期随访的多角度（正/45°斜/90°侧）及动静态照片。
   - **可选输入**：
     - ④ **参考学术论文 (Reference Paper)**：是否有 PDF 论文路径、DOI 或文献文本？（有则走 `evidence-case` 路线；无则直接走 `journey` 纯临床路线，绝不强行编造文献综述）。
2. **模板与视觉设计选择 (`Default 模板` vs `design.md`)**：
   - 询问用户是**使用内置 Default 模板**（默认采用总结自两大标杆案例的 `Ruby Clinical Gallery` 深浅双模态视觉系统与双轨页序），还是**提供自定义的 `design.md`** 来指导 PPT 的色彩、字体与页面构成。
   - 若用户想写 `design.md` 但尚无现成文件，可提供 [design.example.md](assets/design.example.md) 供其参考或按其口述偏好直接生成；若用户需要整理病历素材，可参考 [case-intake-guide.md](assets/case-intake-guide.md)。
3. **汇报场景与讲者信息**：
   - 讲者姓名、单位/机构、会议名称、目标页数（默认单病例无文献约 12–15 页，含文献结合约 14–18 页）及保密/合规声明要求。

---

## 2. 建立临床事实、文献依据与照片映射 (Fact & Photo Mapping)

在开始排版前，务必阅读 [clinical-photo-and-evidence-rules.md](references/clinical-photo-and-evidence-rules.md) 并执行核查：

1. **运行只读照片盘点与缩略图生成**：
   ```bash
   python3 "<skill-dir>/scripts/inventory_images.py" "<photo-folder>" \
     --out "<workspace>/case_work/inventory" \
     --contact-sheet \
     --normalize-copies
   ```
   - 检视生成的 `contact-sheet-*.jpg` 与 `photo-map.csv`，将每张照片绑定到具体的**时点**（如 `术前基线`、`1st Tx 术后即刻`、`1st Tx 后 6 个月 / 2nd Tx 术前`、`2nd Tx 后 3 个月`）、**视角**（`front` / `oblique` / `profile`）、**表情**（`rest` / `smile` / `upgaze`）与**角色**（`clinical` / `treatment_mapping` / `paper_figure`）。
   - **严禁事项**：不得将 EXIF 文件修改时间直接当成治疗日期；不得将第二次治疗前的照片当成原始基线；不得对患者照片做拉伸、镜像翻转或美颜修图。
2. **核对治疗剂量与侧别**：
   - 区分“计划方案”与“实际执行”，表格中 `Left` / `Right` 一律代表患者自身左右侧。
   - 严禁把论文队列平均剂量（如论文均值 0.35 mL）或模板示例剂量套用到未知剂量的患者身上。
3. **提取参考论文（如有）**：
   - 通读 PDF 正文、图例与表格，提取核心解剖结构（如 AM 弓状缘、TTL 泪沟韧带、ORL 眼轮匝肌支持韧带、DMCFP 深层内侧颊脂肪垫、眶隔前间隙）、分型标准（如 Innie vs Outie、Hirmand I–III）、操作步骤、流变学要求及方法学局限，明确“论文做法”与“本例实际操作”的异同。

---

## 3. 确定叙事大纲与美学规范 (Outline & Aesthetic Strategy)

阅读 [default-templates-and-aesthetics.md](references/default-templates-and-aesthetics.md)：

1. **选择叙事轨线**：
   - **无参考论文 → `journey` 模式**：封面 → 全层评估框架（可选 1 页） → 患者基线档案与 FAS/分型评估 → 术前多角度影像 → 首次治疗布点与左右侧剂量表 → 正/斜/侧位同角度前后对比 → 阶段随访与二次治疗方案（如有） → Patient Journey 横向多时点时光走廊 → 临床心得与证据局限 → 讨论与致谢。
   - **有参考论文 → `evidence-case` 模式**：在临床病例前后有机嵌入“临床痛点（如动态香肠样畸形） → 文献核心论断与分层解剖机制（红/蓝/绿/金语义高亮） → 分型与决策树 → 两步/分层操作与流变学匹配 → 本例验证 → 文献方法学局限与本土人群适配讨论”。
2. **应用视觉系统**：
   - **默认模板 (`Ruby Clinical Gallery`)**：临床照片对比页、封面、章节页与 Patient Journey 走廊采用深邃暗色画布（`#111216`）消除白底眩光、凸显面部光影与凹陷变化；文献证据评价、患者文字档案与多维分析卡片采用雅致浅灰白画布（`#F5F5F7`），搭配宝石红（`#C85A70`）、香槟金（`#C5A880`）、韧带蓝（`#4EA3FF`）与安全间隙绿（`#36D47E`）。
   - **自定义 `design.md`**：若用户提供 `design.md`，读取其中的 YAML 令牌、字体、页序与版式要求覆盖默认设置。
   - **空白布局库**：内置的 [default-template.pptx](assets/default-template.pptx) 包含 10 种标准可编辑版式，可供预览或作为母版参考。
3. **自适应裁剪（严禁凑页造假）**：
   - 仅保留有真实数据支撑的页面：单次治疗不生成“二期微调”页；无微笑照不生成“静态 vs 动态微笑”矩阵页；存在关键信息缺口时先向用户展示简洁页序大纲与待确认点。

---

## 4. 构建 PPTX 与逐页视觉质检 (Execution & Visual QA)

阅读 [production-and-schema.md](references/production-and-schema.md)，生成结构化 `slides.json` 并调用内置构建引擎：

```bash
python3 "<skill-dir>/scripts/build_case_deck.py" \
  --config "<workspace>/case_work/slides.json" \
  --out "<output-dir>/Aesthetic_Case_Report.pptx"
```
（若使用自定义 `design.md`，加上 `--design "<path-to-design.md>"`）。

### 强制视觉验证闭环 (Mandatory Visual Verification)
生成 PPTX 后，必须将其渲染为图片并逐页检视，不得仅凭脚本返回码宣称完成：
```bash
soffice --headless --convert-to pdf "<output-dir>/Aesthetic_Case_Report.pptx" --outdir "<workspace>/case_work/qa"
pdftoppm -jpeg -r 150 "<workspace>/case_work/qa/Aesthetic_Case_Report.pdf" "<workspace>/case_work/qa/slide"
```
使用看图工具检查生成的每一页 `slide-*.jpg`：
- **版式与文字**：无文字截断、无重叠溢出、对比度清晰、表格列宽与折行美观。
- **影像保真度**：所有患者照片严格保持原始宽高比（无拉伸变形），术前在左、术后在右，同角度同表情对齐，胶囊标签时点准确。
- **医学严谨度**：未把术后即刻外观夸大为长期疗效，未把静态照片当成动态证据，未把“未记录不良事件”写成“零不良事件”。

确认所有页面达到高美学与医学标准后，向用户交付最终 `.pptx` 文件及简明页面导览。
