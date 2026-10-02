"""Compatibility shim for Python 3.13+.

The stdlib imghdr module was removed in Python 3.13, but this PaddleOCR
version still imports it. This small subset implements imghdr.what().
"""


def what(file, h=None):
    if h is None:
        if isinstance(file, (str, bytes, bytearray)):
            with open(file, "rb") as image_file:
                h = image_file.read(32)
        else:
            location = file.tell()
            h = file.read(32)
            file.seek(location)

    if h.startswith(b"\xff\xd8"):
        return "jpeg"
    if h.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if h[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if h.startswith(b"BM"):
        return "bmp"
    if h.startswith(b"RIFF") and h[8:12] == b"WEBP":
        return "webp"
    if h[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    return None
