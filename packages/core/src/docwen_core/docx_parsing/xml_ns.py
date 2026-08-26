"""OOXML namespace constants used by DOCX->Markdown converters."""

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_V = "urn:schemas-microsoft-com:vml"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

NSMAP_W = {"w": NS_W}
NSMAP_M = {"m": NS_M, "w": NS_W}
NSMAP_WP = {"wp": NS_WP, "a": NS_A, "r": NS_R, "w": NS_W}
NSMAP_WPV = {
    "w": NS_W,
    "wp": NS_WP,
    "a": NS_A,
    "r": NS_R,
    "v": NS_V,
    "pic": NS_PIC,
    "mc": NS_MC,
    "wps": NS_WPS,
}
