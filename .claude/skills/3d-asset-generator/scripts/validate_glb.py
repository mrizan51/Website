#!/usr/bin/env python3
"""Validate a .glb file's structure and print a summary.

Usage: python3 validate_glb.py <file.glb>

Checks the GLB container (header, chunk layout), the glTF JSON (required
fields), that every accessor fits inside its buffer view and buffer, that
POSITION min/max match the binary data, and that all indices are in range.
Exits 0 on success, 1 on any failure.
"""

import json
import struct
import sys

COMPONENT_SIZE = {5120: 1, 5121: 1, 5122: 2, 5123: 2, 5125: 4, 5126: 4}
TYPE_COUNT = {"SCALAR": 1, "VEC2": 2, "VEC3": 3, "VEC4": 4,
              "MAT2": 4, "MAT3": 9, "MAT4": 16}

errors = []


def check(ok, message):
    if not ok:
        errors.append(message)
    return ok


def accessor_bytes(doc, blob, accessor):
    """Return the raw bytes an accessor covers (assumes tightly packed)."""
    view = doc["bufferViews"][accessor["bufferView"]]
    start = view.get("byteOffset", 0) + accessor.get("byteOffset", 0)
    size = (COMPONENT_SIZE[accessor["componentType"]]
            * TYPE_COUNT[accessor["type"]] * accessor["count"])
    return blob[start:start + size]


def main(path):
    data = open(path, "rb").read()
    if not check(len(data) >= 28, "file too small to be a GLB"):
        return

    magic, version, total = struct.unpack_from("<III", data, 0)
    check(magic == 0x46546C67, f"bad magic {magic:#x} (not 'glTF')")
    check(version == 2, f"unsupported glTF version {version}")
    check(total == len(data), f"header length {total} != file size {len(data)}")

    json_len, json_type = struct.unpack_from("<II", data, 12)
    check(json_type == 0x4E4F534A, "first chunk is not JSON")
    check(json_len % 4 == 0, "JSON chunk not 4-byte aligned")
    doc = json.loads(data[20:20 + json_len])

    blob = b""
    bin_offset = 20 + json_len
    if bin_offset < len(data):
        bin_len, bin_type = struct.unpack_from("<II", data, bin_offset)
        check(bin_type == 0x004E4942, "second chunk is not BIN")
        blob = data[bin_offset + 8: bin_offset + 8 + bin_len]
        check(len(blob) == bin_len, "BIN chunk truncated")

    check(doc.get("asset", {}).get("version") == "2.0", "asset.version != 2.0")
    for key in ("scenes", "nodes", "meshes", "accessors", "bufferViews", "buffers"):
        check(key in doc, f"missing top-level '{key}'")
    if errors:
        return

    buf_len = doc["buffers"][0]["byteLength"]
    check(buf_len <= len(blob), f"buffer byteLength {buf_len} > BIN size {len(blob)}")

    for i, view in enumerate(doc["bufferViews"]):
        end = view.get("byteOffset", 0) + view["byteLength"]
        check(end <= buf_len, f"bufferView {i} overruns buffer ({end} > {buf_len})")

    for i, acc in enumerate(doc["accessors"]):
        view = doc["bufferViews"][acc["bufferView"]]
        need = (COMPONENT_SIZE[acc["componentType"]]
                * TYPE_COUNT[acc["type"]] * acc["count"])
        check(acc.get("byteOffset", 0) + need <= view["byteLength"],
              f"accessor {i} overruns its bufferView")

    print(f"{path}: {len(doc['meshes'])} meshes, "
          f"{len(doc.get('materials', []))} materials, "
          f"{len(data)} bytes")
    for mesh in doc["meshes"]:
        for prim in mesh["primitives"]:
            pos = doc["accessors"][prim["attributes"]["POSITION"]]
            check("min" in pos and "max" in pos,
                  f"{mesh['name']}: POSITION accessor missing min/max")
            raw = accessor_bytes(doc, blob, pos)
            floats = struct.unpack(f"<{pos['count'] * 3}f", raw)
            for axis in range(3):
                col = floats[axis::3]
                check(abs(min(col) - pos["min"][axis]) < 1e-4
                      and abs(max(col) - pos["max"][axis]) < 1e-4,
                      f"{mesh['name']}: POSITION min/max mismatch on axis {axis}")
            tri_note = ""
            if "indices" in prim:
                idx_acc = doc["accessors"][prim["indices"]]
                fmt = {5123: "H", 5125: "I"}[idx_acc["componentType"]]
                idx = struct.unpack(f"<{idx_acc['count']}{fmt}",
                                    accessor_bytes(doc, blob, idx_acc))
                check(max(idx) < pos["count"],
                      f"{mesh['name']}: index {max(idx)} out of range")
                tri_note = f", {len(idx) // 3} tris"
            print(f"  - {mesh['name']}: {pos['count']} verts{tri_note}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    try:
        main(sys.argv[1])
    except Exception as exc:  # malformed files should fail, not crash
        errors.append(f"exception while parsing: {exc!r}")
    if errors:
        print(f"INVALID: {sys.argv[1]}")
        for e in errors:
            print(f"  ! {e}")
        sys.exit(1)
    print("VALID")
