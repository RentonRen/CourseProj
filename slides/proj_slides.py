"""Course-project section. Starts with setup and the handover env config."""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

from pathlib import Path

from build_rl_section import INK, MUTED, NAVY, add_muli_run
from sim_slides import _box, _contain_picture, _emu, _fresh, _title

NOTE = "Notebook: hri_handover_tutorial.ipynb"
ROOT = Path(__file__).resolve().parent
PROJ = ROOT / "images" / "proj"


def _section(prs, exemplar: int, title: str, part: str) -> None:
    master = prs.slides[exemplar].slide_layout.slide_master
    layout = None
    for candidate in master.slide_layouts:
        c_sld = candidate._element.find(qn("p:cSld"))
        if c_sld is not None and c_sld.get("name") == "Section Header":
            layout = candidate
            break
    if layout is None:
        raise SystemExit("No light Section Header layout.")
    slide = prs.slides.add_slide(layout)
    for shape in slide.placeholders:
        kind = shape.placeholder_format.type
        if kind == PP_PLACEHOLDER.TITLE:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], title, "Muli ExtraBold", 32, True)
        elif kind == PP_PLACEHOLDER.BODY:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], part, "Muli", 18, False)


def _note(slide) -> None:
    _box(slide, 1.35, 5.12, 5.85, 0.28, [NOTE], size=9, color=MUTED, align=PP_ALIGN.RIGHT)


def _panel(slide, x, y, w, h) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x), _emu(y), _emu(w), _emu(h))
    shape.adjustments[0] = 0.06
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFA)
    shape.line.color.rgb = RGBColor(0xD5, 0xDC, 0xE3)
    shape.line.width = Pt(1)


def _code(slide, x, y, w, h, lines, *, size=13) -> None:
    shape = slide.shapes.add_textbox(_emu(x), _emu(y), _emu(w), _emu(h))
    tf = shape.text_frame
    tf.word_wrap = False
    tf.auto_size = None
    tf._txBody.find(qn("a:bodyPr")).set("anchor", "t")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(3)
        add_muli_run(p, line, "Consolas", size, False, INK)
    return shape


def _setup(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Part 0 — Setup")
    _box(slide, 0.42, 0.76, 4.70, 0.26, [("Hardware", True, 15, NAVY)])
    _box(
        slide,
        0.42,
        0.98,
        4.70,
        0.72,
        [
            [("•  macOS, Linux, or Windows", False, 13, INK)],
            [("•  8 GB RAM, a few GB of disk", False, 13, INK)],
            [("•  GPU optional; CPU is enough to open the notebook", False, 13, INK)],
        ],
    )
    _box(slide, 0.42, 1.70, 4.70, 0.26, [("Get the repo", True, 15, NAVY)])
    _box(
        slide,
        0.42,
        1.92,
        4.70,
        0.90,
        [
            [("•  Clone CourseProj from GitHub", False, 13, INK)],
            [("•  Conda, Python 3.10", False, 13, INK)],
            [("•  pip installs local robosuite", False, 13, INK)],
            [("•  Then open the notebook", False, 13, INK)],
        ],
    )
    _panel(slide, 0.42, 2.86, 4.70, 1.62)
    _code(
        slide,
        0.52,
        2.94,
        4.50,
        1.48,
        [
            "git clone https://github.com/RentonRen/CourseProj.git",
            "cd CourseProj",
            "conda create -n courseproj python=3.10",
            "conda activate courseproj",
            "pip install -r requirements.txt",
            "jupyter notebook hri_handover_tutorial.ipynb",
        ],
        size=11,
    )
    _contain_picture(slide, PROJ / "part0_imports.png", 5.22, 0.80, 4.36, 3.70)
    _box(
        slide,
        5.22,
        4.54,
        4.36,
        0.22,
        ["Part 0 of hri_handover_tutorial.ipynb"],
        size=10,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )
    _box(
        slide,
        0.42,
        4.52,
        4.70,
        0.36,
        [
            "Windows: create C:\\tmp, copy mujoco.dll into robosuite/utils,",
            "set MUJOCO_GL=wgl.",
        ],
        size=10,
        color=MUTED,
    )
    _note(slide)


def _robosuite(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Part 1")
    _box(
        slide,
        0.42,
        0.82,
        4.85,
        3.20,
        [
            [("A robotics library built on MuJoCo.", False, 16, INK)],
            [("It provides:", False, 16, INK)],
            [("•  Ready-made robot arms", False, 16, INK)],
            [("      Panda, XArm7, UR5e, …", False, 15, INK)],
            [("•  Ready-made manipulation tasks", False, 16, INK)],
            [("      Lift, Pick-and-Place, Assembly, …", False, 15, INK)],
            [("•  Cameras, sensors, and a standard", False, 16, INK)],
            [("      observation / action interface", False, 15, INK)],
            [("•  Built-in scenes to test as-is,", False, 16, INK)],
            [("      or change a little and retest", False, 15, INK)],
        ],
    )
    _panel(slide, 5.40, 1.15, 4.18, 1.45)
    _code(
        slide,
        5.52,
        1.30,
        3.95,
        1.20,
        [
            "obs = env.reset()",
            "obs, reward, done, info =",
            "    env.step(action)",
        ],
        size=13,
    )
    _box(
        slide,
        0.42,
        4.78,
        6.2,
        0.22,
        ["Zhu et al., robosuite. arXiv:2009.12293, 2020. 2017 Stanford SVL; now ARISE."],
        size=10,
        color=MUTED,
    )
    _note(slide)


def _env_config(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Part 1")
    _panel(slide, 0.40, 0.88, 5.35, 3.95)
    _code(
        slide,
        0.55,
        1.00,
        5.10,
        3.75,
        [
            "env = suite.make(",
            '    env_name="HumanRobotHandover",',
            '    robots="XArm7Handover",',
            "    has_renderer=False,",
            "    has_offscreen_renderer=True,",
            "    use_camera_obs=True,",
            "    horizon=300,",
            "    control_freq=20,",
            "    reward_shaping=False,",
            "    initialization_noise=None,",
            "    seed=0,",
            "    hard_reset=False,",
            ")",
            "env.start_holding = True",
        ],
        size=13,
    )
    _box(
        slide,
        5.90,
        0.92,
        3.70,
        3.95,
        [
            [("•  One XArm7, hammer in hand", False, 16, INK)],
            [("•  20 Hz, 300 steps, 15 s", False, 16, INK)],
            [("•  Headless images in memory", False, 16, INK)],
            [("•  Cameras on for the tour", False, 16, INK)],
            [("•  Sparse reward at first", False, 16, INK)],
            [("•  seed=0, held start pose", False, 16, INK)],
            [("•  Later: cameras off,", False, 16, INK)],
            [("      reward_shaping=True", False, 16, INK)],
        ],
    )
    _note(slide)


def add_project_slides(prs, exemplar: int) -> None:
    _section(prs, exemplar, "Course project", "Part 3 of 3")
    _setup(prs, exemplar)
    _robosuite(prs, exemplar)
    _env_config(prs, exemplar)
