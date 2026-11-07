#!/usr/bin/env python3
"""
High-aesthetic PPTX builder for aesthetic-case-report-ruby.
Supports:
- Dual-mode gallery aesthetics (Dark Clinical Gallery #111216 + Light Evidence Slate #F5F5F7)
- Custom design.md YAML token overrides (--design design.md)
- Blank reusable layout template generation (--template)
- Aspect-ratio-preserving image containment (no stretching or distortion)
- 10 specialized layouts: cover, section, cards, anatomy, narrative, profile,
  gallery, treatment, comparison, matrix, journey, evidence, closing
"""

import argparse
import json
import os
import re
from pathlib import Path
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

SLIDE_W = 13.333
SLIDE_H = 7.500

DEFAULT_THEME = {
    "style_name": "Ruby Clinical Gallery",
    "bg_dark": "#111216",
    "bg_card_dark": "#1C1D24",
    "border_dark": "#2E303D",
    "bg_light": "#F5F5F7",
    "bg_card_light": "#FFFFFF",
    "border_light": "#D2D2D7",
    "text_on_dark": "#F5F5F7",
    "muted_on_dark": "#A1A1A6",
    "text_on_light": "#1D1D1F",
    "muted_on_light": "#6E6E73",
    "accent_primary": "#C85A70",
    "accent_secondary": "#C5A880",
    "accent_blue": "#4EA3FF",
    "accent_emerald": "#36D47E",
    "font_cn": "PingFang SC",
    "font_en": "Helvetica Neue",
}

DEFAULT_DISCLAIMER_COVER = "（非公开）学术会议交流资料，仅供医疗专业人士参考，请勿外传。"
DEFAULT_DISCLAIMER_CASE = (
    "以上为临床个案分享，仅供本次学术讨论；不同医疗器械/药品的具体使用，"
    "应参见国家药品监督管理局获批的中文产品说明书及相关临床诊疗指南。"
)


def hex_to_rgb(hex_str, fallback="#111216"):
    s = str(hex_str or fallback).strip().lstrip("#")
    if len(s) != 6:
        s = fallback.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def parse_design_md(design_path):
    """Extract theme tokens from a design.md YAML block or key-value lines."""
    text = Path(design_path).read_text(encoding="utf-8")
    overrides = {}
    yaml_match = re.search(r"```(?:yaml|yml)\s*(.*?)```", text, re.DOTALL)
    block = yaml_match.group(1) if yaml_match else text
    for line in block.splitlines():
        m = re.match(r"^\s*([a-zA-Z0-9_]+)\s*:\s*[\"']?([^\"'#\n]+)[\"']?", line)
        if m:
            k, v = m.group(1).strip(), m.group(2).strip()
            if k in DEFAULT_THEME:
                overrides[k] = v
            elif k == "primary_color":
                overrides["accent_primary"] = v
            elif k == "secondary_color":
                overrides["accent_blue"] = v
            elif k == "accent_color":
                overrides["accent_secondary"] = v
            elif k == "font_family_cn":
                overrides["font_cn"] = v.split(",")[0].strip()
            elif k == "font_family_en":
                overrides["font_en"] = v.split(",")[0].strip()
    return overrides


class RubyDeckBuilder:
    def __init__(self, theme=None, base_dir=None, template_mode=False):
        self.theme = {**DEFAULT_THEME, **(theme or {})}
        self.base_dir = Path(base_dir or os.getcwd()).resolve()
        self.template_mode = template_mode
        self.prs = Presentation()
        self.prs.slide_width = Inches(SLIDE_W)
        self.prs.slide_height = Inches(SLIDE_H)
        self.blank = self.prs.slide_layouts[6]

    def c(self, token):
        return hex_to_rgb(self.theme.get(token, DEFAULT_THEME.get(token, "#111216")))

    def set_bg(self, slide, is_dark=True):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = self.c("bg_dark" if is_dark else "bg_light")

    def add_box(self, slide, left, top, width, height, bg_rgb=None, border_rgb=None, rounded=True):
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        sh = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
        if rounded and hasattr(sh, "adjustments") and len(sh.adjustments) > 0:
            sh.adjustments[0] = 0.02
        if bg_rgb is not None:
            sh.fill.solid()
            sh.fill.fore_color.rgb = bg_rgb
        else:
            sh.fill.background()
        if border_rgb is not None:
            sh.line.color.rgb = border_rgb
            sh.line.width = Pt(1)
        else:
            sh.line.fill.background()
        return sh

    def add_text(
        self, slide, text, left, top, width, height,
        size=13, color_rgb=None, bold=False, align=PP_ALIGN.LEFT,
        is_en=False, v_anchor=MSO_ANCHOR.TOP, line_spacing=1.25
    ):
        tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = v_anchor
        tf.margin_left = Inches(0.02)
        tf.margin_right = Inches(0.02)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)

        lines = str(text or "").split("\n")
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = line
            p.alignment = align
            p.line_spacing = line_spacing
            p.space_after = Pt(4) if len(lines) > 1 else Pt(0)
            p.font.name = self.theme["font_en"] if is_en else self.theme["font_cn"]
            p.font.size = Pt(size)
            p.font.bold = bold
            p.font.color.rgb = color_rgb or self.c("text_on_dark")
        return tb

    def add_header(self, slide, title, eyebrow="", is_dark=True):
        eyebrow_color = self.c("accent_primary" if is_dark else "accent_primary")
        title_color = self.c("text_on_dark" if is_dark else "text_on_light")
        if eyebrow:
            self.add_text(
                slide, eyebrow.upper(), 0.75, 0.38, 10.5, 0.28,
                size=10, color_rgb=eyebrow_color, bold=True, is_en=True
            )
            self.add_text(
                slide, title, 0.75, 0.66, 11.8, 0.55,
                size=23, color_rgb=title_color, bold=True
            )
        else:
            self.add_text(
                slide, title, 0.75, 0.48, 11.8, 0.58,
                size=24, color_rgb=title_color, bold=True
            )

    def add_footer(self, slide, citation="", slide_num=1, is_dark=True, notes=""):
        muted = self.c("muted_on_dark" if is_dark else "muted_on_light")
        if citation:
            self.add_text(slide, citation, 0.75, 7.02, 11.0, 0.32, size=8.5, color_rgb=muted)
        self.add_text(
            slide, f"{slide_num:02d}", 11.95, 6.98, 0.63, 0.32,
            size=10, color_rgb=muted, align=PP_ALIGN.RIGHT, is_en=True
        )
        if notes:
            slide.notes_slide.notes_text_frame.text = str(notes)

    def resolve_img_path(self, raw_path):
        if not raw_path:
            return None
        p = Path(raw_path)
        if not p.is_absolute():
            p = (self.base_dir / p).resolve()
        return p if p.is_file() else None

    def add_contained_photo(
        self, slide, img_info, left, top, width, height,
        is_dark=True, pill_color_token=None
    ):
        """
        Draws a framed photo container and fits the image inside with aspect ratio preserved (contain).
        Draws a clean caption pill below the photo.
        """
        if isinstance(img_info, str):
            img_info = {"path": img_info, "label": ""}
        raw_path = img_info.get("path", "")
        label = img_info.get("label", "")
        sublabel = img_info.get("sublabel", "") or img_info.get("note", "")

        card_bg = self.c("bg_card_dark" if is_dark else "bg_card_light")
        border_c = self.c("border_dark" if is_dark else "border_light")
        self.add_box(slide, left, top, width, height, bg_rgb=card_bg, border_rgb=border_c, rounded=True)

        img_file = self.resolve_img_path(raw_path)
        pad = 0.06
        box_w = max(0.2, width - 2 * pad)
        box_h = max(0.2, height - 2 * pad)

        if img_file:
            try:
                with Image.open(img_file) as im:
                    im_fixed = ImageOps.exif_transpose(im)
                    iw, ih = im_fixed.size
                img_ratio = iw / float(ih) if ih else 1.0
                box_ratio = box_w / float(box_h)
                if img_ratio > box_ratio:
                    draw_w = box_w
                    draw_h = box_w / img_ratio
                else:
                    draw_h = box_h
                    draw_w = box_h * img_ratio
                draw_x = left + pad + (box_w - draw_w) / 2.0
                draw_y = top + pad + (box_h - draw_h) / 2.0
                slide.shapes.add_picture(
                    str(img_file), Inches(draw_x), Inches(draw_y),
                    width=Inches(draw_w), height=Inches(draw_h)
                )
            except Exception as exc:
                self._draw_placeholder_text(slide, left, top, width, height, label or f"Load error: {exc}", is_dark)
        else:
            self._draw_placeholder_text(slide, left, top, width, height, label or "待替换临床照片", is_dark)

        if label or sublabel:
            pill_h = 0.56 if sublabel else 0.36
            pill_top = top + height + 0.08
            pill_border = self.c(pill_color_token) if pill_color_token else border_c
            self.add_box(slide, left, pill_top, width, pill_h, bg_rgb=card_bg, border_rgb=pill_border, rounded=True)
            lbl_color = self.c(pill_color_token) if pill_color_token else self.c("text_on_dark" if is_dark else "text_on_light")
            if sublabel:
                self.add_text(
                    slide, label, left + 0.08, pill_top + 0.04, width - 0.16, 0.24,
                    size=10.5, color_rgb=lbl_color, bold=True, align=PP_ALIGN.CENTER
                )
                self.add_text(
                    slide, sublabel, left + 0.08, pill_top + 0.27, width - 0.16, 0.25,
                    size=8.8, color_rgb=self.c("muted_on_dark" if is_dark else "muted_on_light"),
                    align=PP_ALIGN.CENTER
                )
            else:
                self.add_text(
                    slide, label, left + 0.08, pill_top + 0.05, width - 0.16, 0.26,
                    size=10.5, color_rgb=lbl_color, bold=True, align=PP_ALIGN.CENTER,
                    v_anchor=MSO_ANCHOR.MIDDLE
                )

    def _draw_placeholder_text(self, slide, left, top, width, height, label, is_dark):
        muted = self.c("muted_on_dark" if is_dark else "muted_on_light")
        self.add_text(
            slide, f"[ 照片区 / Photo Slot ]\n{label}",
            left + 0.15, top + height / 2.0 - 0.35, width - 0.3, 0.7,
            size=11, color_rgb=muted, align=PP_ALIGN.CENTER, v_anchor=MSO_ANCHOR.MIDDLE
        )

    # -------------------------------------------------------------------------
    # Slide Layout Builders
    # -------------------------------------------------------------------------
    def build_cover(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)

        # Left vertical ruby + champagne dual accent bars
        self.add_box(slide, 0.75, 1.95, 0.10, 2.95, bg_rgb=self.c("accent_primary"), rounded=False)
        self.add_box(slide, 0.91, 2.55, 0.04, 2.35, bg_rgb=self.c("accent_secondary"), rounded=False)

        eyebrow = s.get("eyebrow", "AESTHETIC CLINICAL CASE REPORT · 医美临床个案精研")
        self.add_text(slide, eyebrow, 1.18, 1.92, 10.8, 0.35, size=11, color_rgb=self.c("accent_primary"), bold=True, is_en=True)
        self.add_text(slide, s.get("title", "医美临床病例汇报"), 1.18, 2.32, 11.0, 1.45, size=36, color_rgb=self.c("text_on_dark"), bold=True)
        if s.get("subtitle"):
            self.add_text(slide, s["subtitle"], 1.18, 3.85, 10.8, 0.65, size=18, color_rgb=self.c("accent_secondary"))

        speaker_line = "  |  ".join(filter(None, [s.get("speaker", ""), s.get("institution", ""), s.get("date", "")]))
        if speaker_line or s.get("body"):
            self.add_text(
                slide, speaker_line or s.get("body", ""), 1.18, 5.15, 10.5, 0.65,
                size=13.5, color_rgb=self.c("muted_on_dark")
            )

        self.add_footer(slide, s.get("citation", default_footer or DEFAULT_DISCLAIMER_COVER), idx, is_dark=True, notes=s.get("notes", ""))

    def build_section(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)
        sec_num = str(s.get("section_num", s.get("eyebrow", "01"))).zfill(2)
        self.add_text(slide, sec_num, 0.95, 2.10, 3.0, 1.1, size=60, color_rgb=self.c("accent_primary"), bold=True, is_en=True)
        self.add_box(slide, 0.95, 3.30, 1.2, 0.05, bg_rgb=self.c("accent_secondary"), rounded=False)
        self.add_text(slide, s.get("title", ""), 0.95, 3.55, 10.8, 0.9, size=30, color_rgb=self.c("text_on_dark"), bold=True)
        if s.get("subtitle") or s.get("body"):
            self.add_text(
                slide, s.get("subtitle") or s.get("body"), 0.95, 4.45, 10.5, 0.6,
                size=15, color_rgb=self.c("muted_on_dark"), is_en=True
            )
        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=True, notes=s.get("notes", ""))

    def build_cards(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        is_dark = bool(s.get("dark", False))
        self.set_bg(slide, is_dark=is_dark)
        self.add_header(slide, s.get("title", ""), s.get("eyebrow", ""), is_dark=is_dark)

        cards = s.get("cards", [])
        n = max(1, min(4, len(cards)))
        gap = 0.28
        total_w = 11.833
        card_w = (total_w - gap * (n - 1)) / n
        top = 1.48
        card_h = 5.15

        accent_cycle = ["accent_primary", "accent_blue", "accent_secondary", "accent_emerald"]
        card_bg = self.c("bg_card_dark" if is_dark else "bg_card_light")
        border_c = self.c("border_dark" if is_dark else "border_light")
        txt_main = self.c("text_on_dark" if is_dark else "text_on_light")

        for i, card in enumerate(cards[:4]):
            left = 0.75 + i * (card_w + gap)
            self.add_box(slide, left, top, card_w, card_h, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
            acc = self.c(card.get("accent", accent_cycle[i % len(accent_cycle)]))
            self.add_box(slide, left, top, card_w, 0.08, bg_rgb=acc, rounded=False)

            if card.get("badge"):
                self.add_text(slide, str(card["badge"]), left + 0.24, top + 0.22, card_w - 0.48, 0.28, size=10.5, color_rgb=acc, bold=True, is_en=True)
                title_top = top + 0.50
            else:
                title_top = top + 0.25

            self.add_text(slide, card.get("title", ""), left + 0.24, title_top, card_w - 0.48, 0.55, size=15.5, color_rgb=txt_main, bold=True)
            points = card.get("points", [])
            if isinstance(points, str):
                points = [p.strip() for p in points.split("\n") if p.strip()]
            bullet_str = "\n".join(f"• {pt}" for pt in points)
            self.add_text(
                slide, bullet_str or card.get("body", ""),
                left + 0.24, title_top + 0.62, card_w - 0.48, card_h - (title_top - top) - 0.80,
                size=12, color_rgb=txt_main, line_spacing=1.32
            )

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=is_dark, notes=s.get("notes", ""))

    def build_narrative_or_anatomy(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        is_dark = bool(s.get("dark", True))
        self.set_bg(slide, is_dark=is_dark)
        self.add_header(slide, s.get("title", ""), s.get("eyebrow", ""), is_dark=is_dark)

        images = s.get("images", [])
        top = 1.45
        h = 4.85
        txt_main = self.c("text_on_dark" if is_dark else "text_on_light")
        card_bg = self.c("bg_card_dark" if is_dark else "bg_card_light")
        border_c = self.c("border_dark" if is_dark else "border_light")

        if images or self.template_mode:
            img_entry = images[0] if images else {"path": "", "label": "解剖切面示意图 / 文献核心插图"}
            self.add_contained_photo(slide, img_entry, 0.75, top, 5.55, h, is_dark=is_dark)
            right_x, right_w = 6.55, 6.03
        else:
            right_x, right_w = 0.75, 11.833

        self.add_box(slide, right_x, top, right_w, h + 0.42, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
        items = s.get("items", [])
        if items:
            cur_y = top + 0.24
            step_h = (h + 0.1) / max(1, len(items))
            semantic_colors = ["accent_primary", "accent_blue", "accent_emerald", "accent_secondary"]
            for i, it in enumerate(items[:4]):
                col = self.c(it.get("color", semantic_colors[i % len(semantic_colors)]))
                self.add_box(slide, right_x + 0.25, cur_y + 0.05, 0.08, step_h - 0.18, bg_rgb=col, rounded=False)
                self.add_text(slide, it.get("label", ""), right_x + 0.44, cur_y, right_w - 0.68, 0.30, size=13.5, color_rgb=col, bold=True)
                self.add_text(slide, it.get("desc", ""), right_x + 0.44, cur_y + 0.30, right_w - 0.68, step_h - 0.38, size=11.5, color_rgb=txt_main, line_spacing=1.28)
                cur_y += step_h
        else:
            body_text = s.get("body", "")
            if isinstance(body_text, list):
                body_text = "\n".join(f"• {b}" for b in body_text)
            self.add_text(slide, body_text, right_x + 0.30, top + 0.28, right_w - 0.60, h, size=13, color_rgb=txt_main, line_spacing=1.35)

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=is_dark, notes=s.get("notes", ""))

    def build_profile(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        is_dark = bool(s.get("dark", False))
        self.set_bg(slide, is_dark=is_dark)
        self.add_header(slide, s.get("title", "患者基线档案与术前评估"), s.get("eyebrow", "PATIENT PROFILE & ASSESSMENT"), is_dark=is_dark)

        top = 1.48
        h = 5.15
        col_w = 5.76
        card_bg = self.c("bg_card_dark" if is_dark else "bg_card_light")
        border_c = self.c("border_dark" if is_dark else "border_light")
        txt_main = self.c("text_on_dark" if is_dark else "text_on_light")
        muted = self.c("muted_on_dark" if is_dark else "muted_on_light")

        # Left Card: Patient Demographics & Complaints
        self.add_box(slide, 0.75, top, col_w, h, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
        self.add_box(slide, 0.75, top, col_w, 0.08, bg_rgb=self.c("accent_primary"), rounded=False)
        self.add_text(slide, s.get("left_title", "求美者基本特征与就诊主诉"), 1.02, top + 0.22, col_w - 0.54, 0.38, size=15.5, color_rgb=txt_main, bold=True)

        prof = s.get("profile", {})
        fields = [
            ("病例代号 / 年龄", prof.get("patient_id", "待补充")),
            ("核心主诉 (Chief Complaints)", prof.get("complaints", "待补充")),
            ("既往治疗史 (History)", prof.get("history", "无特殊记录")),
            ("求美目标 (Expectations)", prof.get("expectations", "自然紧致过渡，动态无异常隆起")),
        ]
        cy = top + 0.72
        for lbl, val in fields:
            self.add_text(slide, lbl, 1.02, cy, col_w - 0.54, 0.24, size=11, color_rgb=self.c("accent_primary"), bold=True)
            self.add_text(slide, str(val), 1.02, cy + 0.24, col_w - 0.54, 0.68, size=11.5, color_rgb=txt_main, line_spacing=1.25)
            cy += 1.02

        # Right Card: Pre-treatment Clinical / FAS Assessment
        rx = 6.82
        self.add_box(slide, rx, top, col_w, h, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
        self.add_box(slide, rx, top, col_w, 0.08, bg_rgb=self.c("accent_secondary"), rounded=False)
        self.add_text(slide, s.get("right_title", "术前临床分型与分区评估 (Assessment)"), rx + 0.27, top + 0.22, col_w - 0.54, 0.38, size=15.5, color_rgb=txt_main, bold=True)

        assessments = s.get("assessments", [])
        cy = top + 0.72
        for item in assessments[:5]:
            area = item.get("area", "")
            grade = item.get("grade", "")
            detail = item.get("detail", "")
            self.add_text(slide, f"{area}  |  {grade}" if grade else area, rx + 0.27, cy, col_w - 0.54, 0.24, size=11.5, color_rgb=self.c("accent_blue" if is_dark else "accent_primary"), bold=True)
            if detail:
                self.add_text(slide, detail, rx + 0.27, cy + 0.24, col_w - 0.54, 0.52, size=11, color_rgb=txt_main)
            cy += 0.80

        if s.get("summary_note"):
            self.add_text(slide, f"评估结论：{s['summary_note']}", rx + 0.27, top + h - 0.78, col_w - 0.54, 0.65, size=10.5, color_rgb=muted)

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=is_dark, notes=s.get("notes", ""))

    def build_gallery(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)
        self.add_header(slide, s.get("title", "术前多角度基线影像"), s.get("eyebrow", "BASELINE PHOTOGRAPHY"), is_dark=True)

        images = s.get("images", [])
        if not images and self.template_mode:
            images = [
                {"path": "", "label": "术前 · 正位静态 (Frontal)"},
                {"path": "", "label": "术前 · 45° 斜位 (Oblique)"},
                {"path": "", "label": "术前 · 90° 侧位 (Profile)"},
            ]
        n = max(1, min(4, len(images)))
        gap = 0.26
        total_w = 11.833
        w = (total_w - gap * (n - 1)) / n
        top = 1.45
        h = 4.65
        for i, im in enumerate(images[:4]):
            self.add_contained_photo(slide, im, 0.75 + i * (w + gap), top, w, h, is_dark=True)

        if s.get("body"):
            self.add_text(slide, s["body"], 0.75, 6.58, 11.833, 0.32, size=11, color_rgb=self.c("muted_on_dark"))

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=True, notes=s.get("notes", ""))

    def build_treatment(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        is_dark = bool(s.get("dark", True))
        self.set_bg(slide, is_dark=is_dark)
        self.add_header(slide, s.get("title", "实际治疗方案与精细剂量记录"), s.get("eyebrow", "TREATMENT PLAN & DOSAGE MAPPING"), is_dark=is_dark)

        images = s.get("images", [])
        mapping_img = images[0] if images else ({"path": "", "label": "术前画线布点图 / 进针通道定位"} if self.template_mode else None)

        top = 1.45
        h = 4.70
        if mapping_img:
            self.add_contained_photo(slide, mapping_img, 0.75, top, 3.85, h, is_dark=is_dark, pill_color_token="accent_secondary")
            tx_left, tx_w = 4.82, 7.76
        else:
            tx_left, tx_w = 0.75, 11.833

        headers = s.get("headers", ["解剖靶区 (Site)", "注射层次与工具", "左侧剂量 (Left)", "右侧剂量 (Right)", "产品与流变学特性"])
        rows = s.get("rows", [])
        if not rows and self.template_mode:
            rows = [
                ["Zygomatic arch (颧弓)", "Layer 5 骨膜上 · 27G 锐针", "按记录填写", "按记录填写", "高 G' 强支撑材料"],
                ["Maxilla / DMCFP (深颊)", "深层间隙 · 25G 50mm 钝针", "按记录填写", "按记录填写", "容积支撑型材料"],
                ["Tear Trough (泪沟融合)", "眶隔前间隙 · 25G/27G 钝针", "按记录填写", "按记录填写", "低吸水率泪沟适应证产品"],
            ]

        total_rows = len(rows) + 1
        cols = len(headers)
        row_h = min(0.76, (h + 0.15) / max(2, total_rows))
        tbl_shape = slide.shapes.add_table(total_rows, cols, Inches(tx_left), Inches(top), Inches(tx_w), Inches(row_h * total_rows))
        table = tbl_shape.table

        header_bg = self.c("accent_navy" if is_dark else "accent_primary")
        row_even = self.c("bg_card_dark" if is_dark else "bg_card_light")
        row_odd = RGBColor(23, 25, 32) if is_dark else RGBColor(240, 242, 246)
        txt_color = self.c("text_on_dark" if is_dark else "text_on_light")

        for c_idx, h_text in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = str(h_text)
            p.font.name = self.theme["font_cn"]
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        for r_idx, r_vals in enumerate(rows):
            for c_idx in range(cols):
                val = r_vals[c_idx] if c_idx < len(r_vals) else ""
                cell = table.cell(r_idx + 1, c_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = row_even if r_idx % 2 == 0 else row_odd
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.text = str(val)
                p.font.name = self.theme["font_cn"]
                p.font.size = Pt(10.5)
                p.font.bold = (c_idx == 0)
                p.font.color.rgb = self.c("accent_secondary") if (c_idx == 0 and is_dark) else txt_color

        if s.get("body"):
            self.add_text(
                slide, s["body"], tx_left, top + row_h * total_rows + 0.15, tx_w, 0.50,
                size=10.5, color_rgb=self.c("muted_on_dark" if is_dark else "muted_on_light")
            )

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=is_dark, notes=s.get("notes", ""))

    def build_comparison(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)
        self.add_header(slide, s.get("title", "治疗前后同角度对比"), s.get("eyebrow", "BEFORE & AFTER COMPARISON"), is_dark=True)

        images = s.get("images", [])
        if len(images) < 2 and self.template_mode:
            images = [
                {"path": "", "label": "治疗前基线 (Before · Baseline)"},
                {"path": "", "label": "治疗后随访 (After · Follow-up)"},
            ]

        has_notes_panel = bool(s.get("observations") or s.get("body"))
        top = 1.45
        h = 4.65
        if has_notes_panel:
            img_w = 4.22
            gap = 0.22
            self.add_contained_photo(slide, images[0], 0.75, top, img_w, h, is_dark=True)
            self.add_contained_photo(slide, images[1], 0.75 + img_w + gap, top, img_w, h, is_dark=True, pill_color_token="accent_emerald")

            nx = 0.75 + (img_w + gap) * 2
            nw = 12.583 - nx
            self.add_box(slide, nx, top, nw, h + 0.44, bg_rgb=self.c("bg_card_dark"), border_rgb=self.c("border_dark"), rounded=True)
            self.add_box(slide, nx, top, nw, 0.08, bg_rgb=self.c("accent_emerald"), rounded=False)
            self.add_text(slide, s.get("panel_title", "临床影像观察要点"), nx + 0.20, top + 0.22, nw - 0.40, 0.35, size=14, color_rgb=self.c("text_on_dark"), bold=True)

            obs = s.get("observations", [])
            if isinstance(obs, str):
                obs = [line.strip() for line in obs.split("\n") if line.strip()]
            if not obs and s.get("body"):
                obs = [line.strip() for line in str(s["body"]).split("\n") if line.strip()]
            obs_text = "\n".join(f"• {pt.lstrip('•- ')}" for pt in obs)
            self.add_text(slide, obs_text, nx + 0.20, top + 0.68, nw - 0.40, h - 0.40, size=11.5, color_rgb=self.c("text_on_dark"), line_spacing=1.35)
        else:
            img_w = 5.75
            self.add_contained_photo(slide, images[0], 0.75, top, img_w, h, is_dark=True)
            self.add_contained_photo(slide, images[1], 0.75 + img_w + 0.33, top, img_w, h, is_dark=True, pill_color_token="accent_emerald")

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=True, notes=s.get("notes", ""))

    def build_matrix(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)
        self.add_header(slide, s.get("title", "静态与微笑动态表情对比矩阵"), s.get("eyebrow", "STATIC & DYNAMIC MATRIX"), is_dark=True)

        images = s.get("images", [])
        if len(images) < 4 and self.template_mode:
            images = [
                {"path": "", "label": "术前 · 静态 (Pre · At Rest)"},
                {"path": "", "label": "术后 · 静态 (Post · At Rest)"},
                {"path": "", "label": "术前 · 微笑 (Pre · Smiling)"},
                {"path": "", "label": "术后 · 微笑 (Post · Smiling)"},
            ]

        w = 5.75
        h = 2.05
        gap_x = 0.33
        gap_y = 0.58
        start_y = 1.42
        for i in range(min(4, len(images))):
            col = i % 2
            row = i // 2
            x = 0.75 + col * (w + gap_x)
            y = start_y + row * (h + gap_y)
            pill_tok = "accent_emerald" if col == 1 else None
            self.add_contained_photo(slide, images[i], x, y, w, h, is_dark=True, pill_color_token=pill_tok)

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=True, notes=s.get("notes", ""))

    def build_journey(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)
        self.add_header(slide, s.get("title", "Patient Journey · 纵向随访全景走廊"), s.get("eyebrow", "LONGITUDINAL PATIENT JOURNEY"), is_dark=True)

        stages = s.get("images") or s.get("stages") or []
        if not stages and self.template_mode:
            stages = [
                {"path": "", "label": "Before 1st Tx", "sublabel": "原始术前基线"},
                {"path": "", "label": "After 1st Immediate", "sublabel": "首次治疗即刻"},
                {"path": "", "label": "6 Months Later", "sublabel": "二期治疗前评估"},
                {"path": "", "label": "After 2nd Immediate", "sublabel": "二次治疗即刻"},
                {"path": "", "label": "3M Post-2nd Tx", "sublabel": "远期稳定随访"},
            ]

        n = max(2, min(6, len(stages)))
        gap = 0.18
        total_w = 11.833
        col_w = (total_w - gap * (n - 1)) / n
        top = 1.48
        img_h = 4.25

        for i, st in enumerate(stages[:6]):
            left = 0.75 + i * (col_w + gap)
            pill_tok = "accent_secondary" if i > 0 else None
            self.add_contained_photo(slide, st, left, top, col_w, img_h, is_dark=True, pill_color_token=pill_tok)

        if s.get("body"):
            self.add_text(slide, s["body"], 0.75, 6.55, 11.833, 0.35, size=10.5, color_rgb=self.c("muted_on_dark"))

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=True, notes=s.get("notes", ""))

    def build_evidence(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        is_dark = bool(s.get("dark", False))
        self.set_bg(slide, is_dark=is_dark)
        self.add_header(slide, s.get("title", "文献证据评价与本例临床对照"), s.get("eyebrow", "EVIDENCE APPRAISAL & CASE CORRELATION"), is_dark=is_dark)

        top = 1.48
        h = 5.15
        col_w = 5.76
        card_bg = self.c("bg_card_dark" if is_dark else "bg_card_light")
        border_c = self.c("border_dark" if is_dark else "border_light")
        txt_main = self.c("text_on_dark" if is_dark else "text_on_light")

        # Left column
        self.add_box(slide, 0.75, top, col_w, h, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
        self.add_box(slide, 0.75, top, col_w, 0.08, bg_rgb=self.c("accent_blue"), rounded=False)
        self.add_text(slide, s.get("subtitle", "文献研究结果与方法学审视"), 1.02, top + 0.22, col_w - 0.54, 0.38, size=15.5, color_rgb=txt_main, bold=True)
        left_body = s.get("body", "")
        if isinstance(left_body, list):
            left_body = "\n".join(f"• {x}" for x in left_body)
        self.add_text(slide, left_body, 1.02, top + 0.72, col_w - 0.54, h - 0.95, size=12, color_rgb=txt_main, line_spacing=1.35)

        # Right column
        rx = 6.82
        self.add_box(slide, rx, top, col_w, h, bg_rgb=card_bg, border_rgb=border_c, rounded=True)
        self.add_box(slide, rx, top, col_w, 0.08, bg_rgb=self.c("accent_primary"), rounded=False)
        self.add_text(slide, s.get("secondaryTitle", "本例实践对应与临床证据边界"), rx + 0.27, top + 0.22, col_w - 0.54, 0.38, size=15.5, color_rgb=txt_main, bold=True)
        right_body = s.get("secondaryBody", "")
        if isinstance(right_body, list):
            right_body = "\n".join(f"• {x}" for x in right_body)
        self.add_text(slide, right_body, rx + 0.27, top + 0.72, col_w - 0.54, h - 0.95, size=12, color_rgb=txt_main, line_spacing=1.35)

        self.add_footer(slide, s.get("citation", default_footer), idx, is_dark=is_dark, notes=s.get("notes", ""))

    def build_closing(self, s, idx, default_footer):
        slide = self.prs.slides.add_slide(self.blank)
        self.set_bg(slide, is_dark=True)

        self.add_text(slide, s.get("eyebrow", "SUMMARY & DISCUSSION"), 0.95, 1.25, 11.4, 0.32, size=11, color_rgb=self.c("accent_primary"), bold=True, is_en=True)
        self.add_text(slide, s.get("title", "核心临床体会与讨论"), 0.95, 1.62, 11.4, 0.65, size=28, color_rgb=self.c("text_on_dark"), bold=True)

        takeaways = s.get("takeaways", [])
        if takeaways:
            cy = 2.45
            for i, tk in enumerate(takeaways[:4], 1):
                self.add_box(slide, 0.95, cy, 11.4, 0.82, bg_rgb=self.c("bg_card_dark"), border_rgb=self.c("border_dark"), rounded=True)
                self.add_text(slide, f"0{i}", 1.15, cy + 0.18, 0.55, 0.45, size=16, color_rgb=self.c("accent_secondary"), bold=True, is_en=True)
                self.add_text(slide, tk, 1.75, cy + 0.16, 10.3, 0.52, size=12.5, color_rgb=self.c("text_on_dark"))
                cy += 0.96
        elif s.get("body"):
            self.add_box(slide, 0.95, 2.45, 11.4, 3.2, bg_rgb=self.c("bg_card_dark"), border_rgb=self.c("border_dark"), rounded=True)
            self.add_text(slide, s["body"], 1.25, 2.70, 10.8, 2.7, size=14, color_rgb=self.c("text_on_dark"), line_spacing=1.4)

        self.add_text(
            slide, s.get("thank_you", "Thank You · 欢迎讨论"),
            0.95, 6.18, 11.4, 0.48, size=18, color_rgb=self.c("accent_secondary"), bold=True
        )
        self.add_footer(slide, s.get("citation", default_footer or DEFAULT_DISCLAIMER_CASE), idx, is_dark=True, notes=s.get("notes", ""))


def get_template_slides():
    return [
        {
            "layout": "cover",
            "eyebrow": "AESTHETIC CLINICAL CASE REPORT · 医美病例汇报模板",
            "title": "医美临床病例与学术机制汇报",
            "subtitle": "Patient Journey & Evidence-Based Aesthetic Practice",
            "speaker": "讲者姓名 / 职称",
            "institution": "医疗机构 / 学术会议名称",
        },
        {
            "layout": "anatomy",
            "eyebrow": "01 CLINICAL PROBLEM & ANATOMY",
            "title": "临床问题与分层解剖机制（可选文献支撑）",
            "items": [
                {"label": "核心临床问题", "desc": "说明本例要解决的形态痛点（如静态平整 vs 微笑动态条索、深层容量缺失与韧带牵拉）。", "color": "accent_primary"},
                {"label": "关键韧带与骨相标志", "desc": "明确相关解剖边界（如 AM 弓状缘、TTL 泪沟韧带、ORL 眼轮匝肌支持韧带）。", "color": "accent_blue"},
                {"label": "操作间隙与安全平面", "desc": "标示目标注射层次（如 DMCFP 深层脂肪垫、眶隔前间隙、上颌前/颧前间隙）。", "color": "accent_emerald"},
            ],
        },
        {
            "layout": "profile",
            "eyebrow": "02 PATIENT PROFILE & PRE-TREATMENT ASSESSMENT",
            "title": "患者基线档案与术前评估",
            "profile": {
                "patient_id": "病例代号，年龄",
                "complaints": "患者核心主诉与改善部位",
                "history": "既往注射、光电或手术史",
                "expectations": "患者期望达到的静态与动态美学目标",
            },
            "assessments": [
                {"area": "靶区一（如眶周/泪沟）", "grade": "医生评估/分型", "detail": "如 Innie/Outie 分型、Hirmand 分级或 FAS 评分"},
                {"area": "靶区二（如中面部/深颊）", "grade": "医生评估/分型", "detail": "深层骨相及脂肪垫容量状态"},
                {"area": "靶区三（如鼻唇沟/下颏）", "grade": "医生评估/分型", "detail": "轮廓衔接与软组织松垂度"},
            ],
            "summary_note": "根据医生真实病程记录填写评估结论，不自行推测未记录的诊断。",
        },
        {
            "layout": "gallery",
            "eyebrow": "03 BASELINE PHOTOGRAPHY",
            "title": "治疗前多角度基线影像（正位 / 斜位 / 侧位）",
        },
        {
            "layout": "treatment",
            "eyebrow": "04 TREATMENT PLAN & DOSAGE MAPPING",
            "title": "实际治疗记录：面部布点与精细剂量表",
        },
        {
            "layout": "comparison",
            "eyebrow": "05 BEFORE & AFTER COMPARISON",
            "title": "同角度治疗前后对比（正位 / 斜位 / 侧位）",
            "observations": [
                "仅填写有真实照片与病程记录支持的客观观察要点",
                "明确标注拍摄时间节点（如术后即刻、术后 3 个月、术后 6 个月）",
                "保持同视角、同表情、同取景比例，不拉伸或修图",
            ],
        },
        {
            "layout": "matrix",
            "eyebrow": "06 STATIC & DYNAMIC VERIFICATION",
            "title": "静态与微笑动态表情对比矩阵（按需选用）",
        },
        {
            "layout": "journey",
            "eyebrow": "07 LONGITUDINAL PATIENT JOURNEY",
            "title": "Patient Journey · 多时点纵向随访走廊",
        },
        {
            "layout": "evidence",
            "eyebrow": "08 EVIDENCE CORRELATION & LIMITATIONS",
            "title": "文献依据、本例对应关系与证据局限",
            "subtitle": "文献方法与核心结论（如有参考论文）",
            "body": [
                "研究人群、分型标准与两步/分层操作要点",
                "论文报告的结局指标（如 GAIS）与随访时长",
                "论文方法学局限（如单臂无对照、未设盲、缺乏动态终点）",
            ],
            "secondaryTitle": "本例实践对应与安全性记录",
            "secondaryBody": [
                "本例遵循的解剖原则与个体化剂量调整",
                "随访期内的真实结局与患者反馈",
                "不良事件记录状态及影像可比性说明",
            ],
        },
        {
            "layout": "closing",
            "eyebrow": "09 TAKEAWAYS & DISCUSSION",
            "title": "病例核心经验总结与学术讨论",
            "takeaways": [
                "总结受本例真实病程与随访影像支持的核心临床体会",
                "明确解剖层次、产品流变学特性与剂量控制的匹配关系",
                "提出可供同道交流的开放性临床讨论问题",
            ],
        },
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, help="Path to slides JSON config")
    parser.add_argument("--design", type=Path, help="Optional path to custom design.md")
    parser.add_argument("--template", action="store_true", help="Build the blank 10-layout library PPTX")
    parser.add_argument("--out", type=Path, required=True, help="Output PPTX file path")
    args = parser.parse_args()

    if not args.template and not args.config:
        parser.error("Provide --config <slides.json> or --template")

    theme_overrides = {}
    if args.design and args.design.is_file():
        theme_overrides.update(parse_design_md(args.design))

    if args.template:
        cfg = {"footer": DEFAULT_DISCLAIMER_CASE, "slides": get_template_slides()}
        base_dir = Path.cwd()
    else:
        cfg_path = args.config.expanduser().resolve()
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        base_dir = cfg_path.parent
        if isinstance(cfg.get("theme"), dict):
            theme_overrides.update(cfg["theme"])

    builder = RubyDeckBuilder(theme=theme_overrides, base_dir=base_dir, template_mode=args.template)
    default_footer = cfg.get("footer", DEFAULT_DISCLAIMER_CASE)
    slides = cfg.get("slides", [])
    if not slides:
        raise ValueError("slides array must not be empty")

    dispatch = {
        "cover": builder.build_cover,
        "section": builder.build_section,
        "cards": builder.build_cards,
        "anatomy": builder.build_narrative_or_anatomy,
        "narrative": builder.build_narrative_or_anatomy,
        "profile": builder.build_profile,
        "gallery": builder.build_gallery,
        "treatment": builder.build_treatment,
        "comparison": builder.build_comparison,
        "matrix": builder.build_matrix,
        "journey": builder.build_journey,
        "evidence": builder.build_evidence,
        "closing": builder.build_closing,
    }

    for idx, s in enumerate(slides, 1):
        layout = s.get("layout", "narrative").lower()
        if layout not in dispatch:
            raise ValueError(f"Slide {idx}: unsupported layout '{layout}'")
        dispatch[layout](s, idx, default_footer)

    out_path = args.out.expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    builder.prs.save(str(out_path))
    print(json.dumps({"output": str(out_path), "slides": len(slides), "theme": builder.theme["style_name"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
