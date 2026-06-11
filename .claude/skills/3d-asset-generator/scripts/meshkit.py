"""meshkit -- a tiny, dependency-free 3D mesh toolkit (Python stdlib only).

Build assets by composing primitives, transforming them, and adding them to a
Scene, then export web-ready glTF binary (.glb) or Wavefront (.obj + .mtl).

Conventions
-----------
- Right-handed, Y-up coordinates (glTF convention). Units are meters.
- Triangles wind counter-clockwise when seen from outside (outward normals).
- Colors are hex strings ("#e07a5f") or RGB tuples with components in 0..1.
- box / sphere / cylinder / cone / torus are centered at the origin;
  extrude() runs from y=0 to y=height. Use .floor() to rest a part on y=0.

Example
-------
    from meshkit import Scene, box, cylinder, rgb

    scene = Scene("table")
    scene.add(box(1.2, 0.05, 0.8, color="#8a5a3b", name="top").translate(0, 0.75, 0))
    for sx in (-0.5, 0.5):
        for sz in (-0.3, 0.3):
            scene.add(cylinder(0.04, 0.75, color="#6e4527", name="leg")
                      .translate(sx, 0.375, sz))
    scene.export_glb("table.glb")
"""

from __future__ import annotations

import json
import math
import struct

_EPS = 1e-9

# ----------------------------------------------------------------------------
# small vector helpers (3-tuples)
# ----------------------------------------------------------------------------

def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _cross(a, b):
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _normalize(v):
    length = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    if length < _EPS:
        return (0.0, 0.0, 0.0)
    return (v[0] / length, v[1] / length, v[2] / length)


def rgb(color):
    """Normalize a color spec to an (r, g, b) tuple in 0..1.

    Accepts "#rgb", "#rrggbb" (with or without "#"), or a 3-sequence of floats.
    """
    if isinstance(color, str):
        s = color.lstrip("#")
        if len(s) == 3:
            s = "".join(ch * 2 for ch in s)
        if len(s) != 6:
            raise ValueError(f"bad hex color: {color!r}")
        return tuple(int(s[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    r, g, b = color
    return (float(r), float(g), float(b))


# ----------------------------------------------------------------------------
# Mesh
# ----------------------------------------------------------------------------

class Mesh:
    """An indexed triangle mesh with one material.

    Transform methods mutate the mesh and return self, so they chain:
        cone(0.3, 1).rotate_z(90).translate(1, 0.3, 0)
    """

    def __init__(self, vertices=None, faces=None, name="part",
                 color=(0.8, 0.8, 0.8), metallic=0.0, roughness=0.85,
                 opacity=1.0):
        self.vertices = [tuple(map(float, v)) for v in (vertices or [])]
        self.faces = [tuple(f) for f in (faces or [])]
        self.name = name
        self.color = rgb(color)
        self.metallic = float(metallic)
        self.roughness = float(roughness)
        self.opacity = float(opacity)

    # -- basics ---------------------------------------------------------

    def copy(self, name=None):
        m = Mesh(self.vertices, self.faces, name or self.name,
                 self.color, self.metallic, self.roughness, self.opacity)
        return m

    def apply(self, fn):
        """Apply fn((x, y, z)) -> (x, y, z) to every vertex."""
        self.vertices = [tuple(fn(v)) for v in self.vertices]
        return self

    # -- transforms (chainable) ------------------------------------------

    def translate(self, dx, dy, dz):
        return self.apply(lambda v: (v[0] + dx, v[1] + dy, v[2] + dz))

    def scale(self, sx, sy=None, sz=None):
        if sy is None:
            sy = sz = sx
        flips = (sx < 0) + (sy < 0) + (sz < 0)
        self.apply(lambda v: (v[0] * sx, v[1] * sy, v[2] * sz))
        if flips % 2:  # odd number of negative axes turns the mesh inside out
            self.flip_normals()
        return self

    def rotate_x(self, degrees):
        c, s = math.cos(math.radians(degrees)), math.sin(math.radians(degrees))
        return self.apply(lambda v: (v[0], v[1] * c - v[2] * s, v[1] * s + v[2] * c))

    def rotate_y(self, degrees):
        c, s = math.cos(math.radians(degrees)), math.sin(math.radians(degrees))
        return self.apply(lambda v: (v[0] * c + v[2] * s, v[1], -v[0] * s + v[2] * c))

    def rotate_z(self, degrees):
        c, s = math.cos(math.radians(degrees)), math.sin(math.radians(degrees))
        return self.apply(lambda v: (v[0] * c - v[1] * s, v[0] * s + v[1] * c, v[2]))

    def floor(self, y=0.0):
        """Translate so the lowest point sits at the given y (default 0)."""
        min_y = min(v[1] for v in self.vertices)
        return self.translate(0.0, y - min_y, 0.0)

    def center(self):
        """Translate so the bounding-box center sits at the origin."""
        lo, hi = self.bounds()
        return self.translate(*(-(a + b) / 2.0 for a, b in zip(lo, hi)))

    # -- topology ---------------------------------------------------------

    def flip_normals(self):
        """Reverse winding (use if a part renders inside-out)."""
        self.faces = [(a, c, b) for a, b, c in self.faces]
        return self

    def faceted(self):
        """Split shared vertices so every face shades flat.

        Use for a crisp low-poly look (smooth shading is the default).
        """
        verts, faces = [], []
        for a, b, c in self.faces:
            i = len(verts)
            verts += [self.vertices[a], self.vertices[b], self.vertices[c]]
            faces.append((i, i + 1, i + 2))
        self.vertices, self.faces = verts, faces
        return self

    def smooth_normals(self):
        """Area-weighted per-vertex normals (computed, not stored)."""
        acc = [(0.0, 0.0, 0.0)] * len(self.vertices)
        for a, b, c in self.faces:
            n = _cross(_sub(self.vertices[b], self.vertices[a]),
                       _sub(self.vertices[c], self.vertices[a]))
            for i in (a, b, c):
                acc[i] = (acc[i][0] + n[0], acc[i][1] + n[1], acc[i][2] + n[2])
        return [_normalize(n) if _dot(n, n) > _EPS else (0.0, 1.0, 0.0)
                for n in acc]

    # -- measurements -----------------------------------------------------

    def bounds(self):
        xs, ys, zs = zip(*self.vertices)
        return (min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs))

    def volume(self):
        """Signed volume; positive when windings are consistently outward."""
        total = 0.0
        for a, b, c in self.faces:
            total += _dot(self.vertices[a],
                          _cross(self.vertices[b], self.vertices[c]))
        return total / 6.0


def merge(meshes, name="merged", **material):
    """Combine meshes into one Mesh (single material, from kwargs or first mesh)."""
    meshes = list(meshes)
    first = meshes[0]
    out = Mesh(name=name,
               color=material.pop("color", first.color),
               metallic=material.pop("metallic", first.metallic),
               roughness=material.pop("roughness", first.roughness),
               opacity=material.pop("opacity", first.opacity))
    for m in meshes:
        offset = len(out.vertices)
        out.vertices += m.vertices
        out.faces += [(a + offset, b + offset, c + offset) for a, b, c in m.faces]
    return out


# ----------------------------------------------------------------------------
# primitive builders
# ----------------------------------------------------------------------------

def _ring_angle(s, segments):
    t = 2.0 * math.pi * s / segments
    return math.cos(t), -math.sin(t)  # -sin keeps normals outward (Y-up, RH)


def _revolve(rows, segments):
    """Revolve a profile of (radius, y) rows around the Y axis.

    Rows with radius ~ 0 collapse to a single pole vertex. Returns
    (vertices, faces) with outward winding for a bottom-to-top profile.
    """
    verts, rings = [], []
    for r, y in rows:
        if abs(r) < _EPS:
            rings.append([len(verts)])
            verts.append((0.0, float(y), 0.0))
        else:
            ring = []
            for s in range(segments):
                c, sn = _ring_angle(s, segments)
                ring.append(len(verts))
                verts.append((r * c, float(y), r * sn))
            rings.append(ring)

    faces = []
    for lower, upper in zip(rings, rings[1:]):
        if len(lower) == 1 and len(upper) == 1:
            continue
        if len(lower) == 1:  # apex below a ring
            apex = lower[0]
            for s in range(segments):
                faces.append((apex, upper[(s + 1) % segments], upper[s]))
        elif len(upper) == 1:  # apex above a ring
            apex = upper[0]
            for s in range(segments):
                faces.append((lower[s], lower[(s + 1) % segments], apex))
        else:
            for s in range(segments):
                a, b = lower[s], lower[(s + 1) % segments]
                c, d = upper[(s + 1) % segments], upper[s]
                faces += [(a, b, c), (a, c, d)]
    return verts, faces


def _disc(radius, y, segments, up):
    """Filled circle at height y, facing +Y when up else -Y."""
    verts = [(0.0, float(y), 0.0)]
    for s in range(segments):
        c, sn = _ring_angle(s, segments)
        verts.append((radius * c, float(y), radius * sn))
    faces = []
    for s in range(segments):
        a, b = 1 + s, 1 + (s + 1) % segments
        faces.append((0, a, b) if up else (0, b, a))
    return verts, faces


def box(width=1.0, height=1.0, depth=1.0, **kw):
    """Axis-aligned box centered at the origin (flat-shaded edges)."""
    hx, hy, hz = width / 2.0, height / 2.0, depth / 2.0
    # (normal, tangent u, tangent v) with u x v == normal
    sides = [
        ((1, 0, 0), (0, 1, 0), (0, 0, 1)),
        ((-1, 0, 0), (0, 0, 1), (0, 1, 0)),
        ((0, 1, 0), (0, 0, 1), (1, 0, 0)),
        ((0, -1, 0), (1, 0, 0), (0, 0, 1)),
        ((0, 0, 1), (1, 0, 0), (0, 1, 0)),
        ((0, 0, -1), (0, 1, 0), (1, 0, 0)),
    ]
    half = (hx, hy, hz)
    verts, faces = [], []
    for n, u, v in sides:
        base = len(verts)
        for su, sv in ((-1, -1), (1, -1), (1, 1), (-1, 1)):
            verts.append(tuple(
                (n[i] * abs(half[i] * n[i]) if n[i] else 0)
                + su * u[i] * half[i] + sv * v[i] * half[i]
                for i in range(3)))
        faces += [(base, base + 1, base + 2), (base, base + 2, base + 3)]
    return Mesh(verts, faces, **kw)


def plane(width=1.0, depth=1.0, **kw):
    """Single quad in the XZ plane facing +Y, centered at the origin."""
    hx, hz = width / 2.0, depth / 2.0
    verts = [(-hx, 0, -hz), (-hx, 0, hz), (hx, 0, hz), (hx, 0, -hz)]
    return Mesh(verts, [(0, 1, 2), (0, 2, 3)], **kw)


def sphere(radius=0.5, segments=32, rings=16, **kw):
    """UV sphere centered at the origin."""
    profile = []
    for k in range(rings + 1):
        phi = math.pi * k / rings
        profile.append((radius * math.sin(phi), -radius * math.cos(phi)))
    verts, faces = _revolve(profile, segments)
    return Mesh(verts, faces, **kw)


def cylinder(radius=0.5, height=1.0, segments=32, radius_top=None,
             caps=True, **kw):
    """Cylinder centered at the origin (set radius_top to taper; 0 = cone)."""
    rb = radius
    rt = radius if radius_top is None else radius_top
    h2 = height / 2.0
    verts, faces = _revolve([(rb, -h2), (rt, h2)], segments)
    if caps:
        for r, y, up in ((rb, -h2, False), (rt, h2, True)):
            if r > _EPS:
                cv, cf = _disc(r, y, segments, up)
                off = len(verts)
                verts += cv
                faces += [(a + off, b + off, c + off) for a, b, c in cf]
    return Mesh(verts, faces, **kw)


def cone(radius=0.5, height=1.0, segments=32, **kw):
    """Cone centered at the origin, apex up."""
    return cylinder(radius, height, segments, radius_top=0.0, **kw)


def torus(radius=0.5, tube=0.15, segments=32, sides=16, **kw):
    """Torus lying in the XZ plane. radius = ring center, tube = thickness."""
    verts, faces = [], []
    for i in range(segments):
        c, sn = _ring_angle(i, segments)
        for j in range(sides):
            phi = 2.0 * math.pi * j / sides
            d = radius + tube * math.cos(phi)
            verts.append((d * c, tube * math.sin(phi), d * sn))

    def idx(i, j):
        return (i % segments) * sides + (j % sides)

    for i in range(segments):
        for j in range(sides):
            a, b = idx(i, j), idx(i + 1, j)
            c2, d2 = idx(i + 1, j + 1), idx(i, j + 1)
            faces += [(a, b, c2), (a, c2, d2)]
    return Mesh(verts, faces, **kw)


def lathe(profile, segments=32, **kw):
    """Surface of revolution around Y. profile = [(radius, y), ...] bottom→top.

    Closed solids need radius 0 at both ends (or add discs yourself).
    Example vase: lathe([(0, 0), (0.3, 0), (0.42, 0.25), (0.18, 0.7),
                         (0.22, 0.95), (0, 0.95)])
    """
    verts, faces = _revolve([(float(r), float(y)) for r, y in profile], segments)
    return Mesh(verts, faces, **kw)


# -- extrusion ----------------------------------------------------------------

def _shoelace(outline):
    area = 0.0
    for (x0, z0), (x1, z1) in zip(outline, outline[1:] + outline[:1]):
        area += x0 * z1 - x1 * z0
    return area / 2.0


def _point_in_tri(p, a, b, c):
    def side(p1, p2):
        return (p2[0] - p1[0]) * (p[1] - p1[1]) - (p2[1] - p1[1]) * (p[0] - p1[0])
    d1, d2, d3 = side(a, b), side(b, c), side(c, a)
    return d1 >= -_EPS and d2 >= -_EPS and d3 >= -_EPS


def _ear_clip(outline):
    """Triangulate a simple polygon (positive shoelace orientation)."""
    idx = list(range(len(outline)))
    tris = []
    stuck = 0
    while len(idx) > 3 and stuck <= len(idx):
        n = len(idx)
        clipped = False
        for k in range(n):
            i0, i1, i2 = idx[(k - 1) % n], idx[k], idx[(k + 1) % n]
            a, b, c = outline[i0], outline[i1], outline[i2]
            convex = ((b[0] - a[0]) * (c[1] - a[1])
                      - (b[1] - a[1]) * (c[0] - a[0])) > _EPS
            if not convex:
                continue
            if any(_point_in_tri(outline[j], a, b, c)
                   for j in idx if j not in (i0, i1, i2)):
                continue
            tris.append((i0, i1, i2))
            del idx[k]
            clipped = True
            break
        stuck = 0 if clipped else stuck + 1
    if len(idx) == 3:
        tris.append(tuple(idx))
    else:  # degenerate input -- fall back to a fan so we always return a mesh
        tris += [(idx[0], idx[k], idx[k + 1]) for k in range(1, len(idx) - 1)]
    return tris


def extrude(outline, height=1.0, **kw):
    """Extrude a simple 2D polygon [(x, z), ...] from y=0 up to y=height.

    Handles convex and concave outlines (no holes, no self-intersection).
    For an upright silhouette (a fin, a sign), extrude flat then .rotate_x(-90).
    """
    pts = [(float(x), float(z)) for x, z in outline]
    if len(pts) < 3:
        raise ValueError("outline needs at least 3 points")
    if _shoelace(pts) < 0:
        pts.reverse()
    n = len(pts)
    verts = [(x, 0.0, z) for x, z in pts] + [(x, float(height), z) for x, z in pts]
    faces = []
    for k in range(n):
        b0, b1 = k, (k + 1) % n
        t0, t1 = b0 + n, b1 + n
        faces += [(b1, b0, t0), (b1, t0, t1)]
    for i0, i1, i2 in _ear_clip(pts):
        faces.append((i0, i1, i2))            # bottom (faces -Y)
        faces.append((i0 + n, i2 + n, i1 + n))  # top (faces +Y)
    return Mesh(verts, faces, **kw)


# ----------------------------------------------------------------------------
# Scene + exporters
# ----------------------------------------------------------------------------

class Scene:
    """A named collection of Mesh parts that exports as one file."""

    def __init__(self, name="asset"):
        self.name = name
        self.meshes = []

    def add(self, *meshes):
        self.meshes.extend(meshes)
        return self

    def bounds(self):
        los, his = zip(*(m.bounds() for m in self.meshes))
        return (tuple(min(v[i] for v in los) for i in range(3)),
                tuple(max(v[i] for v in his) for i in range(3)))

    def stats(self):
        verts = sum(len(m.vertices) for m in self.meshes)
        tris = sum(len(m.faces) for m in self.meshes)
        lo, hi = self.bounds()
        size = ", ".join(f"{hi[i] - lo[i]:.2f}" for i in range(3))
        return (f"{self.name}: {len(self.meshes)} parts, {verts} vertices, "
                f"{tris} triangles, size ({size}) m")

    # -- glTF binary ------------------------------------------------------

    def export_glb(self, path):
        if not self.meshes:
            raise ValueError("scene has no meshes")
        blob = bytearray()
        buffer_views, accessors, materials, mat_index = [], [], [], {}
        meshes_json, nodes = [], []

        def add_view(data, target):
            while len(blob) % 4:
                blob.append(0)
            buffer_views.append({"buffer": 0, "byteOffset": len(blob),
                                 "byteLength": len(data), "target": target})
            blob.extend(data)
            return len(buffer_views) - 1

        for mesh in self.meshes:
            if not mesh.faces:
                continue
            positions = mesh.vertices
            normals = mesh.smooth_normals()
            flat_pos = [c for v in positions for c in v]
            flat_nrm = [c for v in normals for c in v]
            flat_idx = [i for f in mesh.faces for i in f]

            pv = add_view(struct.pack(f"<{len(flat_pos)}f", *flat_pos), 34962)
            accessors.append({
                "bufferView": pv, "componentType": 5126, "type": "VEC3",
                "count": len(positions),
                "min": [min(v[i] for v in positions) for i in range(3)],
                "max": [max(v[i] for v in positions) for i in range(3)],
            })
            pa = len(accessors) - 1

            nv = add_view(struct.pack(f"<{len(flat_nrm)}f", *flat_nrm), 34962)
            accessors.append({"bufferView": nv, "componentType": 5126,
                              "type": "VEC3", "count": len(normals)})
            na = len(accessors) - 1

            iv = add_view(struct.pack(f"<{len(flat_idx)}I", *flat_idx), 34963)
            accessors.append({"bufferView": iv, "componentType": 5125,
                              "type": "SCALAR", "count": len(flat_idx)})
            ia = len(accessors) - 1

            key = (mesh.color, mesh.metallic, mesh.roughness, mesh.opacity)
            if key not in mat_index:
                mat = {
                    "name": f"mat_{len(materials)}",
                    "doubleSided": True,
                    "pbrMetallicRoughness": {
                        "baseColorFactor": [*mesh.color, mesh.opacity],
                        "metallicFactor": mesh.metallic,
                        "roughnessFactor": mesh.roughness,
                    },
                }
                if mesh.opacity < 1.0:
                    mat["alphaMode"] = "BLEND"
                mat_index[key] = len(materials)
                materials.append(mat)

            meshes_json.append({"name": mesh.name, "primitives": [{
                "attributes": {"POSITION": pa, "NORMAL": na},
                "indices": ia, "material": mat_index[key],
            }]})
            nodes.append({"mesh": len(meshes_json) - 1, "name": mesh.name})

        while len(blob) % 4:
            blob.append(0)
        doc = {
            "asset": {"version": "2.0", "generator": "meshkit"},
            "scene": 0,
            "scenes": [{"name": self.name, "nodes": list(range(len(nodes)))}],
            "nodes": nodes,
            "meshes": meshes_json,
            "materials": materials,
            "bufferViews": buffer_views,
            "accessors": accessors,
            "buffers": [{"byteLength": len(blob)}],
        }
        body = json.dumps(doc, separators=(",", ":")).encode("utf-8")
        body += b" " * ((4 - len(body) % 4) % 4)
        total = 12 + 8 + len(body) + 8 + len(blob)
        with open(path, "wb") as fh:
            fh.write(struct.pack("<III", 0x46546C67, 2, total))
            fh.write(struct.pack("<II", len(body), 0x4E4F534A))  # 'JSON'
            fh.write(body)
            fh.write(struct.pack("<II", len(blob), 0x004E4942))  # 'BIN'
            fh.write(blob)
        return path

    # -- Wavefront OBJ ----------------------------------------------------

    def export_obj(self, path):
        if not self.meshes:
            raise ValueError("scene has no meshes")
        if not path.endswith(".obj"):
            raise ValueError("path must end with .obj")
        mtl_path = path[:-4] + ".mtl"
        mtl_name = mtl_path.rsplit("/", 1)[-1]
        mat_index, mat_lines = {}, []
        out = [f"# {self.name} (meshkit)", f"mtllib {mtl_name}"]
        offset = 1
        for mesh in self.meshes:
            key = (mesh.color, mesh.opacity)
            if key not in mat_index:
                mat_index[key] = f"mat_{len(mat_index)}"
                r, g, b = mesh.color
                mat_lines += [f"newmtl {mat_index[key]}",
                              f"Kd {r:.4f} {g:.4f} {b:.4f}",
                              f"d {mesh.opacity:.4f}", ""]
            out.append(f"o {mesh.name}")
            out.append(f"usemtl {mat_index[key]}")
            normals = mesh.smooth_normals()
            for v in mesh.vertices:
                out.append(f"v {v[0]:.6g} {v[1]:.6g} {v[2]:.6g}")
            for n in normals:
                out.append(f"vn {n[0]:.4f} {n[1]:.4f} {n[2]:.4f}")
            for a, b, c in mesh.faces:
                out.append("f " + " ".join(f"{i + offset}//{i + offset}"
                                           for i in (a, b, c)))
            offset += len(mesh.vertices)
        with open(path, "w") as fh:
            fh.write("\n".join(out) + "\n")
        with open(mtl_path, "w") as fh:
            fh.write("\n".join(mat_lines) + "\n")
        return path
