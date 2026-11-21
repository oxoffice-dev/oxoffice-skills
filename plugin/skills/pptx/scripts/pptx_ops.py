"""The one helper oxoffice 0.1.0 does not offer yet: every shape of a slide, groups included. It goes away when
oxoffice gains `Shapes.walk()` (oxoffice specs/2026-11-20-skill-helpers-in-the-library.md). Import it from a script
run with the Python that has oxoffice (`oxpy`, this folder on PYTHONPATH). From the command line,
`oxoffice pptx inspect --json DECK` gives the same tree (`slide_details`, `shape_details`, `children`).
"""


def walk(shapes, depth=0):
    """(depth, shape) for every shape of `shapes`, the members of each group after the group, in z-order."""
    for sh in shapes:
        yield depth, sh
        if int(sh.shape_type or 0) == 6:
            yield from walk(sh.shapes, depth + 1)
