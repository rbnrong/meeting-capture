"""Rebuild the bundled icon using Python's standard library (not needed to build the EXE)."""
import struct
from pathlib import Path


def rounded(x, y, left, top, right, bottom, radius):
    if not (left <= x < right and top <= y < bottom):
        return False
    cx = max(left + radius, min(x, right - radius))
    cy = max(top + radius, min(y, bottom - radius))
    return (x - cx) ** 2 + (y - cy) ** 2 <= radius ** 2


images = []
for size in (16, 24, 32, 48, 64, 128, 256):
    pixels = bytearray()
    for row in range(size - 1, -1, -1):
        for col in range(size):
            x, y = (col + .5) * 256 / size, (row + .5) * 256 / size
            rgba = (0, 0, 0, 0)
            if rounded(x, y, 8, 8, 248, 248, 50):
                rgba = (36, 95, 217, 255)
            if rounded(x, y, 55, 64, 201, 190, 15) and not rounded(x, y, 68, 77, 188, 177, 6):
                rgba = (255, 255, 255, 255)
            if ((x - 128) ** 2 + (y - 127) ** 2 < 24 ** 2):
                rgba = (255, 255, 255, 255)
            if (39 <= x < 51 and 47 <= y < 96) or (39 <= x < 91 and 47 <= y < 59):
                rgba = (169, 215, 255, 255)
            if (205 <= x < 217 and 158 <= y < 210) or (165 <= x < 217 and 198 <= y < 210):
                rgba = (169, 215, 255, 255)
            r, g, b, a = rgba
            pixels.extend((b, g, r, a))
    mask = bytes(((size + 31) // 32) * 4 * size)
    header = struct.pack('<IIIHHIIIIII', 40, size, size * 2, 1, 32, 0, len(pixels), 0, 0, 0, 0)
    images.append((size, header + pixels + mask))
offset = 6 + 16 * len(images)
entries = []
for size, data in images:
    entries.append(struct.pack('<BBBBHHII', size % 256, size % 256, 0, 0, 1, 32, len(data), offset))
    offset += len(data)
Path(__file__).with_name('app').joinpath('app.ico').write_bytes(
    struct.pack('<HHH', 0, 1, len(images)) + b''.join(entries) + b''.join(data for _, data in images)
)
