# Controlled native Paint bitmap capture

`paint-native.ole` contains the untouched generic Embed Source representation
from Windows Paint copying a controlled 96 x 64 RGBA image on 2026-10-02.
It contains only Paint class metadata and its embedded bitmap; the separate
Object Descriptor containing the local source path is deliberately excluded.
`paint-qt-image.png` is Qt's lossless serialization of the captured QImage.
Its RGBA pixels match Paint's explicit PNG representation. Paint rounds the
source's partially transparent RGB channels; these captured pixels, rather
than the original source's bytes, define the expected ingress result.

The Paint native OLE bitmap is white-composited and padded to a 16-byte
boundary. It establishes standalone bitmap kind and dimensions; it must not
replace the alpha-preserving QImage. Tests check class, unique source stream,
length/padding, bitmap bounds, format and QImage dimensions. Document text,
HTML, WPS and unknown Office OLE retain their existing ingress priority.

These files support deterministic CI regression. They do not prove a repaired
packaged application or native GUI conversion passed acceptance.
