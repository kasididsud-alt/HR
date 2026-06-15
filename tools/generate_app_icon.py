from __future__ import annotations

import struct
import zlib
from pathlib import Path


SIZE = 256


def _chunk(tag: bytes, data: bytes) -> bytes:
    crc = zlib.crc32(tag + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)


def _write_png(path: Path, pixels: list[tuple[int, int, int, int]], size: int) -> bytes:
    raw = bytearray()
    for y in range(size):
        raw.append(0)
        row = pixels[y * size : (y + 1) * size]
        for r, g, b, a in row:
            raw.extend((r, g, b, a))

    png = bytearray(b"\x89PNG\r\n\x1a\n")
    png.extend(
        _chunk(
            b"IHDR",
            struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0),
        )
    )
    png.extend(_chunk(b"IDAT", zlib.compress(bytes(raw), 9)))
    png.extend(_chunk(b"IEND", b""))
    path.write_bytes(bytes(png))
    return bytes(png)


def _write_ico(path: Path, png_bytes: bytes) -> None:
    icon_dir = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", 0, 0, 0, 0, 1, 32, len(png_bytes), 22)
    path.write_bytes(icon_dir + entry + png_bytes)


def _inside_round_rect(x: int, y: int, x0: int, y0: int, x1: int, y1: int, r: int) -> bool:
    if x < x0 or x >= x1 or y < y0 or y >= y1:
        return False
    if x0 + r <= x < x1 - r or y0 + r <= y < y1 - r:
        return True

    corners = (
        (x0 + r, y0 + r),
        (x1 - r - 1, y0 + r),
        (x0 + r, y1 - r - 1),
        (x1 - r - 1, y1 - r - 1),
    )
    cx = x0 + r if x < x0 + r else x1 - r - 1
    cy = y0 + r if y < y0 + r else y1 - r - 1
    dx = x - cx
    dy = y - cy
    return dx * dx + dy * dy <= r * r


def _fill_round_rect(
    pixels: list[tuple[int, int, int, int]],
    x0: int,
    y0: int,
    x1: int,
    y1: int,
    r: int,
    color: tuple[int, int, int, int],
) -> None:
    for y in range(y0, y1):
        for x in range(x0, x1):
            if _inside_round_rect(x, y, x0, y0, x1, y1, r):
                pixels[y * SIZE + x] = color


def _fill_rect(
    pixels: list[tuple[int, int, int, int]],
    x0: int,
    y0: int,
    x1: int,
    y1: int,
    color: tuple[int, int, int, int],
) -> None:
    for y in range(max(y0, 0), min(y1, SIZE)):
        for x in range(max(x0, 0), min(x1, SIZE)):
            pixels[y * SIZE + x] = color


def _fill_circle(
    pixels: list[tuple[int, int, int, int]],
    cx: int,
    cy: int,
    radius: int,
    color: tuple[int, int, int, int],
) -> None:
    r2 = radius * radius
    for y in range(max(cy - radius, 0), min(cy + radius + 1, SIZE)):
        for x in range(max(cx - radius, 0), min(cx + radius + 1, SIZE)):
            dx = x - cx
            dy = y - cy
            if dx * dx + dy * dy <= r2:
                pixels[y * SIZE + x] = color


def build_icon() -> tuple[bytes, bytes]:
    px = [(0, 0, 0, 0)] * (SIZE * SIZE)

    shadow = (15, 23, 42, 45)
    teal = (15, 118, 110, 255)
    teal_light = (45, 212, 191, 255)
    card = (250, 252, 255, 255)
    slate = (71, 85, 105, 255)
    blue = (37, 99, 235, 255)
    green = (22, 163, 74, 255)
    gold = (245, 158, 11, 255)
    white = (255, 255, 255, 255)

    _fill_round_rect(px, 26, 30, 234, 238, 46, shadow)
    _fill_round_rect(px, 18, 18, 226, 226, 44, teal)
    _fill_round_rect(px, 18, 18, 226, 78, 44, teal_light)

    _fill_round_rect(px, 52, 44, 196, 206, 24, card)
    _fill_round_rect(px, 66, 58, 182, 84, 12, teal)

    for y0 in (104, 140, 176):
        for x0, color in ((68, blue), (108, teal_light), (148, green)):
            _fill_round_rect(px, x0, y0, x0 + 28, y0 + 24, 8, color)

    _fill_circle(px, 186, 178, 36, gold)
    _fill_rect(px, 181, 152, 191, 204, white)
    _fill_rect(px, 169, 162, 199, 170, white)
    _fill_rect(px, 169, 186, 199, 194, white)

    png_path = Path(__file__).resolve().parents[1] / "assets" / "salary_calc.png"
    ico_path = png_path.with_suffix(".ico")
    png_path.parent.mkdir(parents=True, exist_ok=True)

    png_bytes = _write_png(png_path, px, SIZE)
    _write_ico(ico_path, png_bytes)
    return png_path.read_bytes(), ico_path.read_bytes()


def main() -> None:
    _, _ = build_icon()
    print("Generated assets/salary_calc.png and assets/salary_calc.ico")


if __name__ == "__main__":
    main()
