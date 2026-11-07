# PPTX 生成引擎、JSON 配置与视觉质检指南 (Production & Schema Guide)

本指南说明如何使用 `scripts/inventory_images.py` 和 `scripts/build_case_deck.py` 将整理好的患者病程、照片映射表、可选论文依据及 `design.md` 渲染为高美学标准的 16:9 可编辑 PPTX，并完成逐页视觉检查。

---

## 一、标准命令工作流 (Command Line Workflow)

### Step 1: 盘点患者照片文件夹并生成缩略联系表 (Contact Sheet)
```bash
python3 "<skill-dir>/scripts/inventory_images.py" "<photo-folder>" \
  --out "<workspace>/case_work/inventory" \
  --contact-sheet \
  --normalize-copies
```
- 输出：
  - `photo-map.csv` 与 `inventory.json`：包含每张图片的 `image_id`（如 `IMG0001`）、宽高比、SHA-256 重复检测结果。
  - `contact-sheet-01.jpg`：带 `image_id` 编号的深色网格缩略图，供检视并确认每张照片的视角（正/斜/侧位）、表情（静态/微笑/向上看）及时间点。
  - `normalized/`：自动纠正手机/相机 EXIF 旋转方向的无损副本。

### Step 2: 构建临床病例 PPTX（支持 Default 模板或自定义 `design.md`）
```bash
# 使用内置 Default 模板（Ruby Clinical Gallery 主题）
python3 "<skill-dir>/scripts/build_case_deck.py" \
  --config "<workspace>/case_work/slides.json" \
  --out "<output-dir>/Aesthetic_Case_Report.pptx"

# 若用户提供了自定义 design.md，通过 --design 传入自动覆盖色彩与字体令牌
python3 "<skill-dir>/scripts/build_case_deck.py" \
  --config "<workspace>/case_work/slides.json" \
  --design "<path-to-design.md>" \
  --out "<output-dir>/Aesthetic_Case_Report.pptx"
```

### Step 3: 生成/重置内置空白布局库模板 (`assets/default-template.pptx`)
```bash
python3 "<skill-dir>/scripts/build_case_deck.py" \
  --template \
  --out "<skill-dir>/assets/default-template.pptx"
```

---

## 二、`slides.json` 结构化配置协议 (Schema Reference)

`slides.json` 中的图片路径支持绝对路径或相对于 `slides.json` 所在目录的相对路径。所有照片均由引擎按 **等比包含（Aspect-Ratio Contain）** 自动计算居中坐标，绝不发生拉伸变形。

```json
{
  "footer": "以上为临床个案分享，仅供本次学术讨论；不同医疗器械/药品的具体使用，应参见国家药监局获批的中文说明书。",
  "theme": {
    "style_name": "Ruby Clinical Gallery",
    "font_cn": "PingFang SC",
    "font_en": "Helvetica Neue"
  },
  "slides": [
    {
      "layout": "cover",
      "eyebrow": "AESTHETIC CLINICAL CASE REPORT · 文献结合临床个案汇报",
      "title": "泪沟与中面部两步融合技术（The Blending Technique）",
      "subtitle": "解剖间隙机制解析与跨年度纵向 Patient Journey 随访实录",
      "speaker": "刘志昱 医生 (Dr. Zhiyu Liu)",
      "institution": "美学临床实践学术研讨会",
      "date": "2026.11",
      "citation": "（非公开）学术会议交流资料，仅供医疗专业人士参考，请勿外传。"
    },
    {
      "layout": "anatomy",
      "eyebrow": "01 ANATOMICAL RATIONALE",
      "title": "解剖机制：弓状缘（AM）、泪沟韧带（TTL/ORL）与三个间隙",
      "images": [
        {
          "path": "photos/anatomy_orl.png",
          "label": "眶周支持韧带复合体与深层间隙（据文献示意）"
        }
      ],
      "items": [
        {
          "label": "AM（弓状缘）与 Innie / Outie 分型",
          "desc": "Innie 型 AM 位于韧带起点后方，二者间存在可利用的眶隔前间隙；Outie 型二者重合，直接浅填易致香肠样畸形。",
          "color": "accent_primary"
        },
        {
          "label": "第一步 · 深层支撑（DMCFP）",
          "desc": "22G/25G 钝针于深层内侧颊脂肪垫（DMCFP）及梨状孔深间隙补充高 G' 玻尿酸，缩短睑颊距离。",
          "color": "accent_blue"
        },
        {
          "label": "第二步 · 穿韧带融合（眶隔前间隙彗星状沉积）",
          "desc": "经同一进针点垂直于泪沟穿越韧带，彗核置于眶隔前间隙，彗尾退针延入上颌前/颧前间隙。",
          "color": "accent_emerald"
        }
      ],
      "citation": "Beut J, Guisantes E, Jelks GW, Jelks EB, Wong CH. Aesthetic Plast Surg 2026. doi:10.1007/s00266-026-06119-0"
    },
    {
      "layout": "profile",
      "eyebrow": "02 PATIENT PROFILE & ASSESSMENT",
      "title": "患者基线档案与术前评估（Pre-treatment Assessment）",
      "profile": {
        "patient_id": "Momo，女，39 岁",
        "complaints": "中面部扁平凹陷、泪沟与睑颧沟明显，伴轻中度鼻唇沟与下颏轻度后缩",
        "history": "既往仅行基础光电嫩肤，无面部深层注射史",
        "expectations": "改善疲态，要求面中紧致饱满且大笑动态自然无条索感"
      },
      "assessments": [
        {"area": "眶周分型判定", "grade": "Innie 型 · Hirmand II 型", "detail": "向上看（Upgaze）时泪沟边界清晰，伴面中部负向量"},
        {"area": "眶下与泪沟区 (FAS)", "grade": "重度凹陷", "detail": "TTL 与 ORL 皮肤附着处凹陷明显"},
        {"area": "中面部与颧弓 (FAS)", "grade": "中重度容量缺失", "detail": "DMCFP 萎缩导致下睑缺乏支撑"},
        {"area": "下面部与颏部 (FAS)", "grade": "轻中度不足", "detail": "鼻唇沟凹陷，颏部前突度略欠"}
      ],
      "summary_note": "符合“先深层骨相/DMCFP 支撑，再行睑颊交界融合，二期完善下面部轮廓”的分阶段路径。"
    },
    {
      "layout": "treatment",
      "eyebrow": "03 1ST TREATMENT PLAN & MAPPING",
      "title": "首次治疗方案：面部画线布点与精细剂量分配",
      "images": [
        {
          "path": "photos/mapping_front.png",
          "label": "模特 Momo，39 岁 · 术前画线布点"
        }
      ],
      "headers": ["解剖靶区 (Target Site)", "注射层次与工具", "左侧 (Left)", "右侧 (Right)", "产品特性 / 备注"],
      "rows": [
        ["Zygomatic arch (颧弓)", "骨膜上深层 · 锐针/钝针", "0.3 mL", "0.4 mL", "高 G' 强支撑型 HA"],
        ["Maxilla / DMCFP (上颌与深颊)", "深层间隙 · 25G 钝针", "0.7 mL", "0.6 mL", "重建面中前突支撑"],
        ["Cheeks (前颊过渡)", "深皮下/潜在间隙 · 钝针", "0.5 mL", "0.5 mL", "平铺衔接中面部"],
        ["NLF (鼻唇沟)", "梨状孔及深皮下", "0.5 mL", "0.5 mL", "减轻鼻翼基底凹陷"],
        ["Chin (颏部塑形)", "骨膜上深层定点", "共 3.0 mL (中线)", "-", "改善下庭比例与侧颜矢量"]
      ],
      "body": "注：表中 Left / Right 均为患者自身左右侧；剂量严格取自临床实际治疗记录。"
    },
    {
      "layout": "comparison",
      "eyebrow": "04 IMMEDIATE POST-OP COMPARISON",
      "title": "首次治疗前后对比：正位平视静态",
      "images": [
        {"path": "photos/baseline_front.png", "label": "Before · 术前基线"},
        {"path": "photos/immediate_front.png", "label": "Immediate After · 术后即刻"}
      ],
      "observations": [
        "中面部 DMCFP 与上颌前壁获得深层支撑后，苹果肌高光点自然上移",
        "睑颊交界过渡平顺，泪沟与睑颧沟台阶感明显减轻",
        "鼻唇沟深度即时浅化，颏部延伸后面部三庭比例更协调",
        "注：当前右侧为术后即刻外观，远期稳固度需结合 6 个月及跨年随访评估"
      ]
    },
    {
      "layout": "journey",
      "eyebrow": "05 LONGITUDINAL PATIENT JOURNEY",
      "title": "Patient Journey · 跨年度多时点随访全景走廊",
      "images": [
        {"path": "photos/baseline_front.png", "label": "Before 1st Tx", "sublabel": "2024 原始术前基线"},
        {"path": "photos/immediate_front.png", "label": "After 1st Immediate", "sublabel": "首次治疗后即刻"},
        {"path": "photos/journey_6m.png", "label": "6 Months Later", "sublabel": "二期治疗前复评"},
        {"path": "photos/journey_stage2.png", "label": "After 2nd Immediate", "sublabel": "二次治疗后即刻"},
        {"path": "photos/journey_stage_14.png", "label": "3M Post-2nd Tx", "sublabel": "二期术后 3 个月随访"}
      ],
      "body": "同一正位视角按真实时间顺序排列，展示从首次深层建构、6 个月组织融合到二期精细维护的完整旅程。"
    },
    {
      "layout": "evidence",
      "eyebrow": "06 EVIDENCE APPRAISAL & CASE REFLECTION",
      "title": "文献依据、本例对应关系与证据局限",
      "subtitle": "Beut 2026 论文核心结论与方法学局限",
      "body": [
        "论文纳入 38 例全部为 Innie 型（Hirmand II/III 型），平均随访 10.4 个月，GAIS 均达 1–2 分。",
        "核心贡献：将“产品置于韧带上方（眶隔前间隙）还是下方”明确为关键变量，提出垂直穿韧带彗星状沉积。",
        "证据局限：单臂研究无平行注射对照组；GAIS 未报告盲法与评价者；缺乏微笑动态客观终点。"
      ],
      "secondaryTitle": "本例临床实践印证与本土化思考",
      "secondaryBody": [
        "东亚人群常伴中面部骨性发育不足（低平骨相），先做颧弓、上颌与 DMCFP 深层支撑的权重更高。",
        "本例采用分期纵向管理（1st Tx + 6M 后 2nd Tx），避免了单次过量填充导致的假面感。",
        "证据边界：本例展示良好的多时点静态维持效果，未来随访建议补充标准化微笑与向上看动态影像。"
      ]
    },
    {
      "layout": "closing",
      "eyebrow": "07 KEY TAKEAWAYS & DISCUSSION",
      "title": "临床核心总结与讨论议题",
      "takeaways": [
        "先支撑后融合：深层骨膜上与 DMCFP 容量恢复后，眶周浅层所需剂量显著减少，安全性更高。",
        "解剖间隙导向：准确辨识 AM 与 TTL/ORL 的空间关系，避开 SOOF 内淋巴管与血管高危区。",
        "纵向全周期管理：以跨年度 Patient Journey 思维替代单次填沟，实现自然的面部抗衰演进。"
      ]
    }
  ]
}
```

---

## 三、视觉验证与交付质检清单 (Visual QA Checklist)

生成 `.pptx` 后，务必将其转换为图片进行逐页视觉检查（而非仅检查文件生成成功）：

```bash
soffice --headless --convert-to pdf "<output.pptx>" --outdir "<qa-dir>"
pdftoppm -jpeg -r 150 "<qa-dir>/<output.pdf>" "<qa-dir>/slide"
```

使用 `view_file` 检查生成的 `slide-01.jpg`、`slide-02.jpg` 等预览图，逐项核查：
1. **无文字截断或重叠**：标题不超 2 行，卡片文字不超出圆角底板，表格内文字折行自然。
2. **照片无变形、无错位**：人脸宽高比自然（无横向压扁或纵向拉长），Before 在左、After 在右，时间标签与病程完全一致。
3. **色彩对比度达标**：深色页（`#111216`）配浅白字（`#F5F5F7`），浅色页（`#F5F5F7`）配深炭黑字（`#1D1D1F`），绝不出现深底黑字或浅底白字。
4. **医学事实零编造**：剂量、侧别、随访时间点、文献引文与实际输入 100% 吻合。
