#!/usr/bin/env python3
"""Render a .glb to .png with a tiny software rasterizer (stdlib only).

Usage: python3 render_preview.py <file.glb> [out.png] [--size 800] [--grid]

Default is a single 3/4 view; --grid renders front / side / top / 3-4 views
in one image. Flat-shaded with a z-buffer — meant for quick visual checks
of proportions, placement, and color, not final quality.
"""

import json
import math
import struct
import sys
import zlib


def load_glb(path):
    """Return a list of (positions, triangles, rgb_color) per primitive."""
    data = open(path, "rb").read()
    magic, _, _ = struct.unpack_from("<III", data, 0)
    assert magic == 0x46546C67, "not a GLB file"
    json_len = struct.unpack_from("<I", data, 12)[0]
    doc = json.loads(data[20:20 + json_len])
    blob = data[20 + json_len + 8:]

    def acc_data(index, fmt_char, width):
        acc = doc["accessors"][index]
        view = doc["bufferViews"][acc["bufferView"]]
        start = view.get("byteOffset", 0) + acc.get("byteOffset", 0)
        count = acc["count"] * width
        raw = blob[start:start + count * struct.calcsize(fmt_char)]
        flat = struct.unpack(f"<{count}{fmt_char}", raw)
        return [flat[i:i + width] for i in range(0, count, width)]

    fmt_for = {5123: "H", 5125: "I", 5126: "f"}
    prims = []
    for mesh in doc["meshes"]:
        for prim in mesh["primitives"]:
            pos = acc_data(prim["attributes"]["POSITION"], "f", 3)
            idx_acc = doc["accessors"][prim["indices"]]
            idx = acc_data(prim["indices"], fmt_for[idx_acc["componentType"]], 1)
            tris = [(idx[i][0], idx[i + 1][0], idx[i + 2][0])
                    for i in range(0, len(idx), 3)]
            color = (0.8, 0.8, 0.8)
            if "material" in prim:
                pbr = doc["materials"][prim["material"]].get(
                    "pbrMetallicRoughness", {})
                color = tuple(pbr.get("baseColorFactor", [0.8] * 4)[:3])
            prims.append((pos, tris, color))
    return prims


def rotate(v, yaw, pitch):
    cy, sy = math.cos(math.radians(yaw)), math.sin(math.radians(yaw))
    cp, sp = math.cos(math.radians(pitch)), math.sin(math.radians(pitch))
    x, y, z = v
    x, z = x * cy + z * sy, -x * sy + z * cy
    y, z = y * cp - z * sp, y * sp + z * cp
    return (x, y, z)


def render(prims, size, yaw, pitch, supersample=2):
    w = h = size * supersample
    view = [([rotate(p, yaw, pitch) for p in pos], tris, col)
            for pos, tris, col in prims]
    xs = [p[0] for pos, _, _ in view for p in pos]
    ys = [p[1] for pos, _, _ in view for p in pos]
    span = max(max(xs) - min(xs), max(ys) - min(ys), 1e-9) * 1.15
    scale = w / span
    cx, cy = (max(xs) + min(xs)) / 2, (max(ys) + min(ys)) / 2

    bg = (22, 25, 31)
    pix = bytearray(bg * (w * h))
    depth = [-1e30] * (w * h)
    light = (0.45, 0.75, 0.49)

    for pos, tris, col in view:
        for a, b, c in tris:
            p0, p1, p2 = pos[a], pos[b], pos[c]
            ux, uy, uz = (p1[0] - p0[0], p1[1] - p0[1], p1[2] - p0[2])
            vx, vy, vz = (p2[0] - p0[0], p2[1] - p0[1], p2[2] - p0[2])
            nx, ny, nz = (uy * vz - uz * vy, uz * vx - ux * vz,
                          ux * vy - uy * vx)
            length = math.sqrt(nx * nx + ny * ny + nz * nz) or 1e-12
            lam = abs(nx * light[0] + ny * light[1] + nz * light[2]) / length
            shade = 0.3 + 0.7 * lam
            rgb = bytes(min(255, int(c2 * shade * 255)) for c2 in col)

            sx0, sy0 = (p0[0] - cx) * scale + w / 2, h / 2 - (p0[1] - cy) * scale
            sx1, sy1 = (p1[0] - cx) * scale + w / 2, h / 2 - (p1[1] - cy) * scale
            sx2, sy2 = (p2[0] - cx) * scale + w / 2, h / 2 - (p2[1] - cy) * scale
            area = (sx1 - sx0) * (sy2 - sy0) - (sy1 - sy0) * (sx2 - sx0)
            if abs(area) < 1e-9:
                continue
            min_x = max(0, int(min(sx0, sx1, sx2)))
            max_x = min(w - 1, int(max(sx0, sx1, sx2)) + 1)
            min_y = max(0, int(min(sy0, sy1, sy2)))
            max_y = min(h - 1, int(max(sy0, sy1, sy2)) + 1)
            for py in range(min_y, max_y + 1):
                for px in range(min_x, max_x + 1):
                    w0 = ((sx1 - px) * (sy2 - py) - (sx2 - px) * (sy1 - py)) / area
                    w1 = ((sx2 - px) * (sy0 - py) - (sx0 - px) * (sy2 - py)) / area
                    w2 = 1 - w0 - w1
                    if w0 < 0 or w1 < 0 or w2 < 0:
                        continue
                    z = w0 * p0[2] + w1 * p1[2] + w2 * p2[2]
                    i = py * w + px
                    if z > depth[i]:
                        depth[i] = z
                        pix[i * 3:i * 3 + 3] = rgb

    # box-filter downsample for cheap anti-aliasing
    out = bytearray(size * size * 3)
    ss = supersample
    for py in range(size):
        for px in range(size):
            for ch in range(3):
                total = 0
                for dy in range(ss):
                    row = (py * ss + dy) * w
                    for dx in range(ss):
                        total += pix[(row + px * ss + dx) * 3 + ch]
                out[(py * size + px) * 3 + ch] = total // (ss * ss)
    return out


def write_png(path, rgb, width, height):
    def chunk(tag, payload):
        body = tag + payload
        return (struct.pack(">I", len(payload)) + body
                + struct.pack(">I", zlib.crc32(body)))

    raw = b"".join(b"\x00" + bytes(rgb[y * width * 3:(y + 1) * width * 3])
                   for y in range(height))
    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw, 6))
           + chunk(b"IEND", b""))
    open(path, "wb").write(png)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(1)
    src = args[0]
    dst = args[1] if len(args) > 1 else src.rsplit(".", 1)[0] + ".png"
    size = 800
    if "--size" in sys.argv:
        size = int(sys.argv[sys.argv.index("--size") + 1])
    prims = load_glb(src)

    if "--grid" in sys.argv:
        half = size // 2
        views = [("front", 0, 0), ("side", 90, 0),
                 ("top", 0, 90), ("3/4", 35, 25)]
        canvas = bytearray(size * size * 3)
        for k, (_, yaw, pitch) in enumerate(views):
            tile = render(prims, half, yaw, pitch)
            ox, oy = (k % 2) * half, (k // 2) * half
            for y in range(half):
                start = ((oy + y) * size + ox) * 3
                canvas[start:start + half * 3] = \
                    tile[y * half * 3:(y + 1) * half * 3]
        write_png(dst, canvas, size, size)
    else:
        write_png(dst, render(prims, size, 35, 25), size, size)
    print(f"wrote {dst}")


if __name__ == "__main__":
    main()
