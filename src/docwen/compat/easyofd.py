from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def apply_easyofd_patches() -> None:
    try:
        from easyofd.draw.draw_pdf import DrawPDF
    except Exception:
        return

    current = getattr(DrawPDF, "draw_annotation", None)
    if getattr(current, "_docwen_patched", False):
        return

    def draw_annotation_patched(self, canvas, annota_info, images, page_size):
        try:
            img_list = []
            if not annota_info:
                return

            for _key, annotation in annota_info.items():
                try:
                    if not annotation:
                        continue

                    anno_type_obj = annotation.get("AnnoType")
                    if not anno_type_obj or anno_type_obj.get("type") != "Stamp":
                        continue

                    img_obj = annotation.get("ImgageObject") or annotation.get("ImageObject")
                    if not img_obj:
                        continue

                    boundary_str = img_obj.get("Boundary") or ""
                    pos_str = boundary_str.split(" ") if boundary_str else []
                    pos = [float(i) for i in pos_str] if pos_str else []

                    appearance = annotation.get("Appearance") or {}
                    wrap_boundary_str = appearance.get("Boundary") or ""
                    wrap_pos_str = wrap_boundary_str.split(" ") if wrap_boundary_str else []
                    wrap_pos = [float(i) for i in wrap_pos_str] if wrap_pos_str else []

                    ctm_str = img_obj.get("CTM") or ""
                    ctm_split = ctm_str.split(" ") if ctm_str else []
                    ctm = [float(i) for i in ctm_split] if ctm_split else []

                    img_list.append(
                        {
                            "wrap_pos": wrap_pos,
                            "pos": pos,
                            "CTM": ctm,
                            "ResourceID": img_obj.get("ResourceID", ""),
                        }
                    )
                except Exception:
                    continue

            if hasattr(self, "draw_img"):
                self.draw_img(canvas, img_list, images, page_size)
        except Exception as e:
            logger.warning("处理OFD注释时发生异常 (Patched): %s", e)

    draw_annotation_patched._docwen_patched = True  # type: ignore[attr-defined]
    DrawPDF.draw_annotation = draw_annotation_patched
