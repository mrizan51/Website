---
name: 3d-asset-generator
description: Generates 3D assets procedurally from text descriptions and exports web-ready GLB/glTF or OBJ files. Use when the user asks to create, generate, or edit a 3D model, mesh, prop, or scene (e.g. "make a low-poly rocket", "generate a 3D mug as GLB", "build a tree model for my Three.js site"). Pure-Python toolkit, no external dependencies.
argument-hint: [description of the 3D asset to generate]
---

# 3D Asset Generator

Build 3D assets by composing primitives with the bundled `scripts/meshkit.py`
(Python stdlib only — never pip install anything for this). Export `.glb`
(preferred, web-ready) or `.obj`, validate the output, and report stats.

## Workflow

1. **Plan the asset.** Break the request into parts, each mapped to a
   primitive (or a lathe/extrude profile). Pick a palette of 3–6 hex colors
   and real-world dimensions in meters (props ≈ 0.3–2 m). State the plan in
   one short paragraph before coding.
2. **Write a build script** named `<asset>_build.py` (working dir is fine;
   don't commit it unless asked) and run it:

   ```python
   import sys
   sys.path.insert(0, "<path to this skill>/scripts")
   from meshkit import *

   scene = Scene("mug")
   scene.add(cylinder(radius=0.045, height=0.1, segments=48,
                      color="#e07a5f", name="body").translate(0, 0.05, 0))
   scene.add(torus(radius=0.032, tube=0.008, segments=32, sides=12,
                   color="#e07a5f", name="handle")
             .rotate_x(90).translate(0.058, 0.05, 0))
   print(scene.stats())
   scene.export_glb("mug.glb")
   ```
3. **Validate:** `python3 <skill>/scripts/validate_glb.py mug.glb` must print
   `VALID`. If it doesn't, fix the build script — never hand-edit the GLB.
4. **Look at it before declaring victory:**
   `python3 <skill>/scripts/render_preview.py mug.glb --grid` renders
   front/side/top/3-4 views to a PNG (stdlib software rasterizer, ~5 s).
   Read the PNG and check proportions, part placement, colors, and that
   nothing is floating or inside-out. Fix the build script and re-render
   until it reads correctly.
5. **Report** the stats line, the output paths, and attach or show the
   preview PNG. For an interactive look, the user can copy
   `references/viewer.html` next to the asset, serve with
   `python3 -m http.server 8000`, and open `viewer.html?file=mug.glb`
   (or just drag the GLB onto the page).
6. **Iterate** on feedback by editing the build script and re-running.

## meshkit API

Primitives (all return a `Mesh`; all accept `name=`, `color=` hex or RGB
tuple, `metallic=` 0–1, `roughness=` 0–1, `opacity=` 0–1):

| Primitive | Notes |
|---|---|
| `box(width, height, depth)` | centered at origin, flat-shaded |
| `sphere(radius, segments=32, rings=16)` | UV sphere, centered |
| `cylinder(radius, height, segments=32, radius_top=None, caps=True)` | `radius_top` tapers (0 = cone, e.g. cups, lampshades) |
| `cone(radius, height, segments=32)` | apex up, centered |
| `torus(radius, tube, segments=32, sides=16)` | flat in XZ plane; `rotate_x(90)` to stand upright |
| `plane(width, depth)` | single quad facing +Y |
| `lathe(profile, segments=32)` | revolve `[(radius, y), ...]` bottom→top around Y — vases, bottles, chess pieces |
| `extrude(outline, height)` | extrude 2D `[(x, z), ...]` polygon from y=0 up; concave OK, no holes |
| `merge(meshes, name=, color=)` | weld parts into one Mesh |

Transforms mutate and chain: `.translate(x, y, z)`, `.rotate_x/y/z(degrees)`,
`.scale(s)` or `.scale(sx, sy, sz)`, `.floor()` (rest on y=0), `.center()`,
`.copy()`, `.flip_normals()`, `.faceted()`.

`Scene(name)`: `.add(*meshes)`, `.stats()`, `.export_glb(path)`,
`.export_obj(path)` (also writes a `.mtl`).

## Conventions and tips

- Y-up, right-handed, units in meters; CCW winding = outward normals.
  If a custom part renders inside-out, call `.flip_normals()`.
- Build each part at the origin, then transform into place. Repeated parts:
  `part.copy().rotate_y(120 * k)`. Let parts interpenetrate slightly rather
  than trying to line up faces exactly — there is no CSG.
- Upright silhouettes (fins, signs, gables): draw the outline in (x, z),
  `extrude(outline, thickness).rotate_x(-90)` makes it vertical.
- Low-poly style: 6–12 segments **and** call `.faceted()` for flat shading.
  Smooth style: 32–48 segments (default shading is smooth).
- Make the ground contact explicit: `.floor()` the lowest parts so the asset
  sits on y=0; keep the whole asset roughly centered on the Y axis.
- Vary roughness for material read: matte plastic ≈ 0.9, wood ≈ 0.8,
  metal ≈ 0.35 with `metallic=1`, glass-ish ≈ `opacity=0.4, roughness=0.15`.
- Keep meshes under ~50k triangles; bump segment counts only where curvature
  is visible.

## Skill files

- `scripts/meshkit.py` — mesh toolkit (primitives, transforms, GLB/OBJ export)
- `scripts/validate_glb.py` — structural validator for generated GLBs
- `scripts/render_preview.py` — GLB → PNG software renderer for self-checks
- `references/examples.md` — worked examples (rocket, low-poly tree) and snippets
- `references/viewer.html` — drag-and-drop Three.js viewer for the browser
