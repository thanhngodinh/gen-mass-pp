"""Church Presentation Liturgical Tools for Giáo xứ Tân Định (9:16 Vertical Ratio).

Specialized tools for auditing slide design against pp_framework.md rules,
exporting high-res preview images, and fast in-place slide patching.
"""

import json
import logging
import os
import sys
from typing import List, Optional

from pydantic import BaseModel, ConfigDict, Field

from backend import ppt
from utils.color import hex_to_int, int_to_hex
from utils.offload import run_offloaded

logger = logging.getLogger(__name__)

# BẢNG MÀU CHUẨN HOÁ THEO PP_FRAMEWORK.MD
CANONICAL_PALETTE = {
    "#CC3300": "Đỏ cam (Tiêu đề bài hát, ĐÁP CA, BÀI ĐỌC, BỘ LỄ)",
    "#C00000": "Đỏ rượu (Ký hiệu ĐK, 1., Lời dẫn, Bìa)",
    "#FF0000": "Đỏ tươi (Số phiên khúc Đáp ca)",
    "#035135": "Xanh rêu đậm (Chân trang Chuyển cảnh)",
    "#000000": "Đen (Lời ca, kinh, câu đáp, trích yếu)",
    "#FFFFFF": "Trắng (Khung hộp chữ trên hoa văn)",
}

CANONICAL_HEXES = set(CANONICAL_PALETTE.keys())

# CHUẨN KÍCH THƯỚC ĐỨNG 9:16
EXPECTED_WIDTH_PT = 405.0   # 5.625 inches
EXPECTED_HEIGHT_PT = 720.0  # 10.000 inches
TOLERANCE_PT = 2.0


# ---------------------------------------------------------------------------
# Pydantic Input Models
# ---------------------------------------------------------------------------
class ChurchAuditSlideInput(BaseModel):
    """Input for auditing a slide against Tân Định liturgical guidelines."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(
        default=1,
        description="1-based index of the slide to audit.",
    )
    file_path: Optional[str] = Field(
        default=None,
        description="Optional full path to .pptx file. If not open, opens it first.",
    )


class ChurchExportPreviewInput(BaseModel):
    """Input for exporting a slide to high-res PNG for visual inspection."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to export.",
    )
    file_path: Optional[str] = Field(
        default=None,
        description="Optional path to .pptx file if not already active.",
    )
    output_dir: Optional[str] = Field(
        default="c:\\PP\\scratch",
        description="Directory to save the preview PNG.",
    )
    width: int = Field(
        default=1080,
        description="Image width in pixels (default 1080 for 9:16).",
    )
    height: int = Field(
        default=1920,
        description="Image height in pixels (default 1920 for 9:16).",
    )


class ChurchQuickPatchInput(BaseModel):
    """Input for fast in-place text and color modification of a slide."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    slide_index: int = Field(
        ...,
        description="1-based index of the slide to patch.",
    )
    search_text: Optional[str] = Field(
        default=None,
        description="Substring to find and replace in text shapes.",
    )
    new_text: Optional[str] = Field(
        default=None,
        description="Replacement text (or new text if shape_index is specified).",
    )
    shape_index: Optional[int] = Field(
        default=None,
        description="Optional 1-based shape index to directly update.",
    )
    font_color_hex: Optional[str] = Field(
        default=None,
        description="Optional hex color to set (e.g. '#CC3300', '#C00000', '#000000').",
    )
    bold: Optional[bool] = Field(
        default=None,
        description="Optional bold setting (True/False).",
    )
    file_path: Optional[str] = Field(
        default=None,
        description="Optional path to .pptx file.",
    )
    save: bool = Field(
        default=True,
        description="Whether to save the presentation after patching.",
    )


# ---------------------------------------------------------------------------
# COM Implementation Helpers
# ---------------------------------------------------------------------------
def _get_active_or_open(file_path: Optional[str] = None):
    app = ppt._get_app_impl(allow_launch=True)
    if file_path:
        norm_path = os.path.abspath(file_path)
        try:
            for pres in app.Presentations:
                if os.path.abspath(pres.FullName).lower() == norm_path.lower():
                    return pres
        except Exception:
            pass
        return app.Presentations.Open(norm_path)
    try:
        return app.ActivePresentation
    except Exception:
        if app.Presentations.Count > 0:
            return app.Presentations(1)
        raise RuntimeError("No presentation is currently open in PowerPoint. Please provide file_path or open a presentation first.")



def _classify_slide_type(all_text: str, shapes_info: list) -> str:
    text_upper = all_text.upper()
    if "BÀI ĐỌC" in text_upper or "ĐÁP CA" in text_upper or "TIN MỪNG" in text_upper or "TUNG HÔ" in text_upper:
        return "reading"
    if "ĐIỆN THOẠI" in text_upper or "RƯỚC LỄ" in text_upper and "ĐIỀU KIỆN" in text_upper:
        return "notice"
    if "GIÁO XỨ TÂN ĐỊNH" in text_upper:
        # Check if title or divider
        if any(sp.get("top_in", 0) > 7.0 for sp in shapes_info):
            return "divider" if any("035135" in sp.get("font_colors", []) for sp in shapes_info) else "title"
    if "KINH" in text_upper or "BỘ LỄ" in text_upper or "THÁNH!" in text_upper or "CHIÊN THIÊN CHÚA" in text_upper:
        return "prayer"
    if "ĐK." in text_upper or "1." in text_upper or "2." in text_upper:
        return "song"
    return "general"


def _church_audit_slide_impl(slide_index: int, file_path: Optional[str] = None) -> str:
    pres = _get_active_or_open(file_path)
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(f"slide_index {slide_index} out of range (1 to {pres.Slides.Count})")

    slide = pres.Slides(slide_index)
    setup = pres.PageSetup
    width_pt = setup.SlideWidth
    height_pt = setup.SlideHeight
    width_in = width_pt / 72.0
    height_in = height_pt / 72.0

    is_9_16 = abs(width_pt - EXPECTED_WIDTH_PT) <= TOLERANCE_PT and abs(height_pt - EXPECTED_HEIGHT_PT) <= TOLERANCE_PT

    shapes_data = []
    all_slide_text = []
    unauthorized_fonts = set()
    unauthorized_colors = set()
    non_bold_runs = 0
    white_boxes = []

    for i in range(1, slide.Shapes.Count + 1):
        sp = slide.Shapes(i)
        left_in = sp.Left / 72.0
        top_in = sp.Top / 72.0
        w_in = sp.Width / 72.0
        h_in = sp.Height / 72.0

        shape_text = ""
        shape_fonts = set()
        shape_colors = set()
        is_bold_all = True

        if sp.HasTextFrame and sp.TextFrame.HasText:
            tr = sp.TextFrame.TextRange
            shape_text = tr.Text
            all_slide_text.append(shape_text)

            # Check lines/words formatting
            try:
                for p_idx in range(1, tr.Paragraphs().Count + 1):
                    p = tr.Paragraphs(p_idx)
                    font = p.Font
                    fn = font.Name
                    shape_fonts.add(fn)
                    if fn.lower() != "arial":
                        unauthorized_fonts.add(fn)

                    if not font.Bold:
                        is_bold_all = False
                        non_bold_runs += 1

                    try:
                        c_int = font.Color.RGB
                        c_hex = int_to_hex(c_int).upper()
                        shape_colors.add(c_hex)
                        if c_hex not in CANONICAL_HEXES:
                            unauthorized_colors.add(c_hex)
                    except Exception:
                        pass
            except Exception:
                pass

        # Check solid white box fill
        is_white_box = False
        try:
            if sp.Fill.Type == 1:  # msoFillSolid
                fill_hex = int_to_hex(sp.Fill.ForeColor.RGB).upper()
                if fill_hex == "#FFFFFF":
                    is_white_box = True
                    white_boxes.append({
                        "shape_index": i,
                        "name": sp.Name,
                        "bounds_in": f"({left_in:.2f}\", {top_in:.2f}\", {w_in:.2f}\" x {h_in:.2f}\")",
                        "top_in": round(top_in, 2),
                        "height_in": round(h_in, 2),
                        "width_in": round(w_in, 2)
                    })
        except Exception:
            pass

        shapes_data.append({
            "shape_index": i,
            "name": sp.Name,
            "bounds": f"({left_in:.2f}\", {top_in:.2f}\", {w_in:.2f}\" x {h_in:.2f}\")",
            "top_in": top_in,
            "is_white_box": is_white_box,
            "text_sample": shape_text.replace("\r", " ").strip()[:60] if shape_text else "",
            "fonts": list(shape_fonts),
            "font_colors": list(shape_colors),
            "is_bold": is_bold_all
        })

    full_text = " ".join(all_slide_text)
    slide_type = _classify_slide_type(full_text, shapes_data)

    warnings = []
    passes = []

    # 1. Dimension check
    if is_9_16:
        passes.append(f"Kích thước chuẩn đứng 9:16 ({width_in:.3f}\" x {height_in:.3f}\").")
    else:
        warnings.append(f"Kích thước KHÔNG ĐẠT chuẩn 9:16 Tân Định (Hiện tại: {width_in:.2f}\" x {height_in:.2f}\", Chuẩn: 5.625\" x 10.000\").")

    # 2. White container check
    if slide_type in ("song", "prayer"):
        if len(white_boxes) >= 2:
            passes.append(f"Slide {slide_type} có đủ {len(white_boxes)} khung hộp trắng tách biệt tiêu đề và nội dung trên nền viền xanh.")
        else:
            warnings.append(f"Slide {slide_type} THIẾU KHUNG HỘP TRẮNG! Yêu cầu 2 hộp trắng (#FFFFFF) trên nền xanh lá #08622A. Hiện chỉ có {len(white_boxes)}.")
    elif slide_type == "reading":
        if len(white_boxes) >= 2:
            passes.append(f"Slide Lời Chúa ({slide_type}) có đủ {len(white_boxes)} khung hộp trắng tách biệt tiêu đề và nội dung.")
        else:
            warnings.append(f"Slide Lời Chúa ({slide_type}) THIẾU KHUNG HỘP TRẮNG! Yêu cầu 2 hộp trắng (#FFFFFF) tại (0.30\", 0.28\") và (0.31\", 2.40\"). Hiện chỉ có {len(white_boxes)}.")
    elif slide_type == "notice":
        if len(white_boxes) >= 1:
            passes.append("Slide Thông báo (notice) có khung thẻ trắng bao quanh chuẩn.")
        else:
            warnings.append("Slide Thông báo (notice) thiếu khung thẻ trắng bao quanh slide (0.31\", 0.40\", 5.00\" x 9.20\").")
    elif slide_type == "divider":
        if any(box["top_in"] >= 7.5 for box in white_boxes):
            passes.append("Slide Chuyển cảnh (divider) có khung hộp trắng chân trang chuẩn.")
        else:
            warnings.append("Slide Chuyển cảnh (divider) thiếu khung hộp trắng chân trang tại (0, 8.38\", 5.625\" x 1.62\").")


    # 3. Typography & Color check
    if unauthorized_fonts:
        warnings.append(f"Phát hiện phông chữ không chuẩn: {list(unauthorized_fonts)}. Quy định bắt buộc: 100% Arial.")
    else:
        passes.append("100% phông chữ tuân thủ chuẩn Arial.")

    if non_bold_runs > 0:
        warnings.append(f"Có {non_bold_runs} đoạn chữ chưa in đậm (Bold = False). Quy định: Toàn bộ chữ phải In Đậm.")
    else:
        passes.append("100% chữ đều In Đậm (Bold = True).")

    if unauthorized_colors:
        warnings.append(f"Phát hiện mã màu ngoài bảng màu chuẩn phụng vụ: {list(unauthorized_colors)}. Các màu hợp lệ: {list(CANONICAL_PALETTE.keys())}")
    else:
        passes.append("Bảng màu chữ hoàn toàn tuân thủ bảng màu phụng vụ Tân Định.")

    verdict = "PASS" if not warnings else ("WARNING" if is_9_16 else "FAIL")

    result = {
        "slide_index": slide_index,
        "total_slides": pres.Slides.Count,
        "detected_type": slide_type,
        "verdict": verdict,
        "dimensions": {
            "width_inches": round(width_in, 3),
            "height_inches": round(height_in, 3),
            "is_9_16": is_9_16
        },
        "white_boxes_count": len(white_boxes),
        "white_boxes": white_boxes,
        "audit_passes": passes,
        "audit_warnings": warnings,
        "shapes": shapes_data
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


def _church_export_preview_impl(
    slide_index: int,
    file_path: Optional[str] = None,
    output_dir: str = "c:\\PP\\scratch",
    width: int = 1080,
    height: int = 1920
) -> str:
    pres = _get_active_or_open(file_path)
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(f"slide_index {slide_index} out of range (1 to {pres.Slides.Count})")

    os.makedirs(output_dir, exist_ok=True)
    slide = pres.Slides(slide_index)
    pres_stem = os.path.splitext(os.path.basename(pres.FullName))[0] if pres.FullName else "deck"
    image_name = f"{pres_stem}_slide_{slide_index:03d}.png"
    target_path = os.path.abspath(os.path.join(output_dir, image_name))

    slide.Export(target_path, "PNG", width, height)

    if not os.path.exists(target_path):
        raise RuntimeError(f"Export failed: file not created at {target_path}")

    size_bytes = os.path.getsize(target_path)
    result = {
        "status": "success",
        "slide_index": slide_index,
        "image_path": target_path,
        "dimensions": f"{width}x{height}",
        "size_bytes": size_bytes,
        "message": f"Slide {slide_index} exported successfully to {target_path}"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


def _church_quick_patch_impl(
    slide_index: int,
    search_text: Optional[str] = None,
    new_text: Optional[str] = None,
    shape_index: Optional[int] = None,
    font_color_hex: Optional[str] = None,
    bold: Optional[bool] = None,
    file_path: Optional[str] = None,
    save: bool = True
) -> str:
    pres = _get_active_or_open(file_path)
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(f"slide_index {slide_index} out of range (1 to {pres.Slides.Count})")

    slide = pres.Slides(slide_index)
    modified_shapes = []

    # Case 1: Targeted shape index
    if shape_index is not None:
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            raise ValueError(f"shape_index {shape_index} out of range (1 to {slide.Shapes.Count})")
        sp = slide.Shapes(shape_index)
        if sp.HasTextFrame:
            tr = sp.TextFrame.TextRange
            if new_text is not None:
                if search_text:
                    tr.Text = tr.Text.replace(search_text, new_text)
                else:
                    tr.Text = new_text
            if font_color_hex:
                tr.Font.Color.RGB = hex_to_int(font_color_hex)
            if bold is not None:
                tr.Font.Bold = -1 if bold else 0
            modified_shapes.append(shape_index)

    # Case 2: Search across all shapes on slide
    elif search_text and new_text is not None:
        for i in range(1, slide.Shapes.Count + 1):
            sp = slide.Shapes(i)
            if sp.HasTextFrame and sp.TextFrame.HasText:
                tr = sp.TextFrame.TextRange
                if search_text in tr.Text:
                    tr.Text = tr.Text.replace(search_text, new_text)
                    if font_color_hex:
                        tr.Font.Color.RGB = hex_to_int(font_color_hex)
                    if bold is not None:
                        tr.Font.Bold = -1 if bold else 0
                    modified_shapes.append(i)

    # Case 3: Only color / bold update across text shapes
    elif font_color_hex or bold is not None:
        for i in range(1, slide.Shapes.Count + 1):
            sp = slide.Shapes(i)
            if sp.HasTextFrame and sp.TextFrame.HasText:
                tr = sp.TextFrame.TextRange
                if font_color_hex:
                    tr.Font.Color.RGB = hex_to_int(font_color_hex)
                if bold is not None:
                    tr.Font.Bold = -1 if bold else 0
                modified_shapes.append(i)

    if save and modified_shapes:
        try:
            pres.Save()
        except Exception as e:
            logger.warning("Could not auto-save presentation: %s", e)

    result = {
        "status": "success" if modified_shapes else "no_match",
        "slide_index": slide_index,
        "modified_shapes_count": len(modified_shapes),
        "modified_shapes": modified_shapes,
        "saved": save and bool(modified_shapes)
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Public Async Entry Points
# ---------------------------------------------------------------------------
async def church_audit_slide(params: ChurchAuditSlideInput) -> str:
    """Audit slide design against Tân Định liturgy rules (dimensions, white containers, colors, Arial bold)."""
    return await run_offloaded(ppt.execute, _church_audit_slide_impl, params.slide_index, params.file_path)


async def church_export_preview(params: ChurchExportPreviewInput) -> str:
    """Export a high-resolution 9:16 PNG image of the slide for instant visual inspection."""
    return await run_offloaded(
        ppt.execute,
        _church_export_preview_impl,
        params.slide_index,
        params.file_path,
        params.output_dir,
        params.width,
        params.height
    )


async def church_quick_patch(params: ChurchQuickPatchInput) -> str:
    """Quickly update text, font colors, or bold styling on a slide in-place without rebuilding all slides."""
    return await run_offloaded(
        ppt.execute,
        _church_quick_patch_impl,
        params.slide_index,
        params.search_text,
        params.new_text,
        params.shape_index,
        params.font_color_hex,
        params.bold,
        params.file_path,
        params.save
    )
