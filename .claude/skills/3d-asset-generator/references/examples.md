# Worked examples

Both scripts assume `sys.path` includes the skill's `scripts/` directory:

```python
import sys
sys.path.insert(0, ".claude/skills/3d-asset-generator/scripts")
from meshkit import *
```

## Example 1 — toy rocket (smooth style)

Plan: white cylinder body, red nose cone, tapered dark nozzle, a red band,
and three fins made from an extruded triangle silhouette rotated upright and
copied around the Y axis.

```python
scene = Scene("rocket")

scene.add(cylinder(radius=0.5, height=2.0, segments=48,
                   color="#e8e4dc", name="body").translate(0, 1.3, 0))
scene.add(cone(radius=0.5, height=0.9, segments=48,
               color="#d64545", name="nose").translate(0, 2.75, 0))
scene.add(torus(radius=0.5, tube=0.06, segments=48, sides=12,
                color="#d64545", name="band").translate(0, 2.2, 0))
scene.add(cylinder(radius=0.32, height=0.35, segments=32, radius_top=0.42,
                   color="#3a3f4a", metallic=0.8, roughness=0.45,
                   name="nozzle").translate(0, 0.18, 0))

# Fin silhouette in (x, z): root edge against the body, swept tip.
fin_outline = [(0.0, 0.0), (0.65, 0.0), (0.65, 0.25), (0.0, 0.95)]
fin = (extrude(fin_outline, height=0.07, color="#d64545", name="fin")
       .rotate_x(-90)            # stand the silhouette upright
       .translate(0.45, 0.05, 0.035))
for k in range(3):
    scene.add(fin.copy(name=f"fin_{k}").rotate_y(120 * k))

print(scene.stats())
scene.export_glb("rocket.glb")
```

## Example 2 — low-poly tree (faceted style)

Plan: tapered trunk plus three stacked, shrinking cones. Low segment counts
and `.faceted()` give the flat-shaded look.

```python
scene = Scene("tree")

scene.add(cylinder(radius=0.16, height=0.7, segments=7, radius_top=0.12,
                   color="#7a5230", name="trunk").faceted().translate(0, 0.35, 0))
for k, (r, h, y) in enumerate([(0.75, 0.9, 1.0), (0.6, 0.8, 1.55),
                               (0.42, 0.7, 2.05)]):
    scene.add(cone(radius=r, height=h, segments=8, color="#3f7d44",
                   name=f"canopy_{k}").faceted().translate(0, y, 0))

print(scene.stats())
scene.export_glb("tree.glb")
scene.export_obj("tree.obj")   # OBJ + MTL variant
```

## Snippet library

```python
# Vase / bottle / chess pawn: lathe a profile, radius 0 at both ends closes it.
vase = lathe([(0, 0), (0.30, 0.0), (0.42, 0.25), (0.18, 0.70),
              (0.22, 0.95), (0, 0.95)], segments=40, color="#5b7d9e")

# Star plate: concave extrude.
import math
star = []
for i in range(10):
    r = 0.5 if i % 2 == 0 else 0.21
    a = math.pi * i / 5
    star.append((r * math.cos(a), r * math.sin(a)))
badge = extrude(star, height=0.08, color="#e3b341", metallic=0.9,
                roughness=0.35, name="star")

# Wheels/rings standing upright:
wheel = torus(radius=0.3, tube=0.08, color="#222").rotate_x(90)

# Ground the finished asset:
lowest = min(m.bounds()[0][1] for m in scene.meshes)
for m in scene.meshes:
    m.translate(0, -lowest, 0)
```
