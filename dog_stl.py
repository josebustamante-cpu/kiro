#!/usr/bin/env python3
"""Generate a simple low-poly dog as a binary STL file.

The dog is built from geometric primitives (boxes and cylinders) that are
positioned in space and merged into a single triangle mesh. Dimensions are
in millimetres so the result is ready for 3D printing.
"""

import math
import struct

# ---------------------------------------------------------------------------
# Mesh helpers
# ---------------------------------------------------------------------------

Triangle = tuple  # (v0, v1, v2) where each v is (x, y, z)
triangles = []


def add_triangle(v0, v1, v2):
    triangles.append((v0, v1, v2))


def add_quad(a, b, c, d):
    """Add a quad (a,b,c,d ordered around the face) as two triangles."""
    add_triangle(a, b, c)
    add_triangle(a, c, d)


def add_box(center, size):
    """Axis-aligned box centred at `center` with full `size` (sx, sy, sz)."""
    cx, cy, cz = center
    sx, sy, sz = size
    hx, hy, hz = sx / 2.0, sy / 2.0, sz / 2.0

    # 8 corners
    p = [
        (cx - hx, cy - hy, cz - hz),  # 0
        (cx + hx, cy - hy, cz - hz),  # 1
        (cx + hx, cy + hy, cz - hz),  # 2
        (cx - hx, cy + hy, cz - hz),  # 3
        (cx - hx, cy - hy, cz + hz),  # 4
        (cx + hx, cy - hy, cz + hz),  # 5
        (cx + hx, cy + hy, cz + hz),  # 6
        (cx - hx, cy + hy, cz + hz),  # 7
    ]

    add_quad(p[0], p[1], p[2], p[3])  # bottom
    add_quad(p[4], p[7], p[6], p[5])  # top
    add_quad(p[0], p[4], p[5], p[1])  # front (-y)
    add_quad(p[2], p[6], p[7], p[3])  # back (+y)
    add_quad(p[1], p[5], p[6], p[2])  # right (+x)
    add_quad(p[3], p[7], p[4], p[0])  # left (-x)


def add_cylinder(base, height, radius, axis="z", segments=24):
    """Cylinder starting at `base`, extruded `height` along `axis`."""
    bx, by, bz = base

    def point(angle, h):
        c = radius * math.cos(angle)
        s = radius * math.sin(angle)
        if axis == "z":
            return (bx + c, by + s, bz + h)
        if axis == "y":
            return (bx + c, by + h, bz + s)
        # axis == "x"
        return (bx + h, by + c, bz + s)

    if axis == "z":
        top_c = (bx, by, bz + height)
        bot_c = (bx, by, bz)
    elif axis == "y":
        top_c = (bx, by + height, bz)
        bot_c = (bx, by, bz)
    else:
        top_c = (bx + height, by, bz)
        bot_c = (bx, by, bz)

    for i in range(segments):
        a0 = 2 * math.pi * i / segments
        a1 = 2 * math.pi * (i + 1) / segments
        b0, b1 = point(a0, 0.0), point(a1, 0.0)
        t0, t1 = point(a0, height), point(a1, height)

        add_quad(b0, t0, t1, b1)        # side wall
        add_triangle(bot_c, b1, b0)     # bottom cap
        add_triangle(top_c, t0, t1)     # top cap


def add_sphere(center, radius, segments=20, rings=14):
    """UV sphere centred at `center`."""
    cx, cy, cz = center

    def point(theta, phi):
        return (
            cx + radius * math.sin(theta) * math.cos(phi),
            cy + radius * math.sin(theta) * math.sin(phi),
            cz + radius * math.cos(theta),
        )

    for r in range(rings):
        t0 = math.pi * r / rings
        t1 = math.pi * (r + 1) / rings
        for s in range(segments):
            p0 = 2 * math.pi * s / segments
            p1 = 2 * math.pi * (s + 1) / segments
            a = point(t0, p0)
            b = point(t1, p0)
            c = point(t1, p1)
            d = point(t0, p1)
            add_quad(a, b, c, d)


# ---------------------------------------------------------------------------
# Build the dog (coordinates in mm).  X = length, Y = width, Z = height.
# ---------------------------------------------------------------------------

# Torso
add_box(center=(0, 0, 45), size=(80, 34, 34))

# Chest / rear rounding
add_sphere(center=(38, 0, 45), radius=18)   # chest front
add_sphere(center=(-40, 0, 45), radius=18)  # hindquarters

# Neck (angled-ish, approximated with a box)
add_box(center=(48, 0, 58), size=(22, 24, 30))

# Head
add_box(center=(64, 0, 70), size=(28, 26, 26))
add_sphere(center=(64, 0, 70), radius=15)

# Snout / muzzle
add_box(center=(82, 0, 64), size=(18, 16, 14))

# Nose
add_sphere(center=(92, 0, 64), radius=5)

# Ears (two flat boxes on top of head)
add_box(center=(58, 11, 86), size=(10, 4, 18))
add_box(center=(58, -11, 86), size=(10, 4, 18))

# Eyes
add_sphere(center=(76, 8, 74), radius=3.2)
add_sphere(center=(76, -8, 74), radius=3.2)

# Legs (4 cylinders going down to the ground at z=0)
leg_r = 6.5
leg_h = 28
for (lx, ly) in [(32, 13), (32, -13), (-32, 13), (-32, -13)]:
    add_cylinder(base=(lx, ly, 0), height=leg_h, radius=leg_r, axis="z")
    add_sphere(center=(lx, ly, 0), radius=leg_r)  # rounded paw

# Tail (cylinder angled up at the back)
add_cylinder(base=(-48, 0, 55), height=30, radius=4, axis="x")  # short stub
add_sphere(center=(-48, 0, 55), radius=5)


# ---------------------------------------------------------------------------
# Geometry utilities + STL writer
# ---------------------------------------------------------------------------

def normal(v0, v1, v2):
    ux, uy, uz = (v1[0] - v0[0], v1[1] - v0[1], v1[2] - v0[2])
    vx, vy, vz = (v2[0] - v0[0], v2[1] - v0[1], v2[2] - v0[2])
    nx = uy * vz - uz * vy
    ny = uz * vx - ux * vz
    nz = ux * vy - uy * vx
    length = math.sqrt(nx * nx + ny * ny + nz * nz)
    if length == 0:
        return (0.0, 0.0, 0.0)
    return (nx / length, ny / length, nz / length)


def write_binary_stl(path, tris):
    with open(path, "wb") as f:
        header = b"Low-poly dog generated by Kiro".ljust(80, b" ")
        f.write(header)
        f.write(struct.pack("<I", len(tris)))
        for v0, v1, v2 in tris:
            nx, ny, nz = normal(v0, v1, v2)
            f.write(struct.pack("<3f", nx, ny, nz))
            for v in (v0, v1, v2):
                f.write(struct.pack("<3f", *v))
            f.write(struct.pack("<H", 0))


if __name__ == "__main__":
    write_binary_stl("perro.stl", triangles)
    print(f"perro.stl written with {len(triangles)} triangles")
