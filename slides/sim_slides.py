"""Simulation section. Text is editable. Pictures and videos are separate files."""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from build_rl_section import INK, MUTED, NAVY, add_muli_run, duplicate_slide

ROOT = Path(__file__).resolve().parent
SIM = ROOT / "images" / "sim"
CREDIT = "Kaifeng Zhang, COMS4733 Lecture 13, Columbia University."

INCH = 914400


def _emu(inches: float) -> Emu:
    return Emu(int(inches * INCH))


def _fresh(prs, exemplar: int):
    slide = duplicate_slide(prs, exemplar)
    tree = slide.shapes._spTree
    for child in list(tree):
        if child.tag not in {qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("p:grpSp"), qn("p:graphicFrame")}:
            continue
        name = child.find(".//" + qn("p:cNvPr"))
        if name is not None and "Slide Number" in (name.get("name") or ""):
            continue
        tree.remove(child)
    return slide


def _box(slide, x, y, w, h, paragraphs, *, size=18, color=INK, align=PP_ALIGN.LEFT, anchor="t"):
    shape = slide.shapes.add_textbox(_emu(x), _emu(y), _emu(w), _emu(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf._txBody.find(qn("a:bodyPr")).set("anchor", anchor)
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(2)
        if isinstance(para, str):
            runs = [(para, False, size, color)]
        elif isinstance(para, tuple):
            runs = [para]
        else:
            runs = para
        for text, bold, psize, pcolor in runs:
            add_muli_run(p, text, "Muli", psize or size, bold, pcolor or color)
    return shape


def _title(slide, text: str) -> None:
    _box(slide, 0.42, 0.16, 9.1, 0.62, [text], size=36, color=INK, align=PP_ALIGN.LEFT)
    # Title weight: rewrite the run as extra bold.
    shape = slide.shapes[-1]
    run = shape.text_frame.paragraphs[0].runs[0]
    run.font.bold = True
    run.font.name = "Muli ExtraBold"


def _credit(slide, extra: str = "") -> None:
    text = extra if extra else CREDIT
    _box(slide, 1.35, 5.12, 5.85, 0.28, [text], size=9, color=MUTED, align=PP_ALIGN.RIGHT)


def _section(prs, exemplar: int) -> None:
    master = prs.slides[exemplar].slide_layout.slide_master
    layout = None
    for candidate in master.slide_layouts:
        c_sld = candidate._element.find(qn("p:cSld"))
        if c_sld is not None and c_sld.get("name") == "Section Header":
            layout = candidate
            break
    if layout is None:
        raise SystemExit("No light Section Header layout.")
    from pptx.enum.shapes import PP_PLACEHOLDER

    slide = prs.slides.add_slide(layout)
    for shape in slide.placeholders:
        kind = shape.placeholder_format.type
        if kind == PP_PLACEHOLDER.TITLE:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], "Simulation", "Muli ExtraBold", 32, True)
        elif kind == PP_PLACEHOLDER.BODY:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], "Part 2 of 3", "Muli", 18, False)


def _font(size: int, bold: bool = False):
    path = "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf"
    return ImageFont.truetype(path, size)


def _make_images() -> None:
    SIM.mkdir(parents=True, exist_ok=True)
    font = _font(28, True)
    small = _font(22)

    def canvas(w, h, color=(255, 255, 255)):
        return Image.new("RGB", (w, h), color)

    # Scalability: many identical arms.
    im = canvas(900, 560, (236, 239, 242))
    d = ImageDraw.Draw(im)
    for row in range(4):
        for col in range(7):
            x, y = 40 + col * 124, 36 + row * 130
            d.rectangle((x, y + 70, x + 90, y + 86), fill=(40, 70, 110))
            d.rectangle((x + 38, y + 28, x + 52, y + 72), fill=(70, 110, 150))
            d.polygon([(x + 45, y + 28), (x + 78, y + 8), (x + 86, y + 20), (x + 52, y + 40)], fill=(20, 50, 90))
    im.save(SIM / "offer_scale.png")

    # Ground truth: real vs sim.
    im = canvas(900, 560)
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, 448, 560), fill=(232, 220, 200))
    d.rectangle((452, 0, 900, 560), fill=(214, 228, 214))
    d.rectangle((70, 250, 250, 430), fill=(180, 90, 60))
    d.rectangle((150, 140, 190, 260), fill=(90, 90, 96))
    d.rectangle((560, 250, 740, 430), fill=(70, 130, 90))
    d.rectangle((630, 140, 670, 260), fill=(40, 70, 50))
    d.text((24, 18), "Real", fill=(40, 40, 40), font=font)
    d.text((476, 18), "Sim", fill=(40, 40, 40), font=font)
    d.text((24, 500), "pose, contact, pixels", fill=(60, 60, 60), font=small)
    im.save(SIM / "offer_truth.png")

    # Bridge design and reality.
    im = canvas(900, 560, (245, 245, 245))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((40, 80, 400, 460), radius=16, outline=(0, 51, 102), width=4)
    d.text((70, 180), "Design", fill=(0, 51, 102), font=font)
    d.text((70, 230), "in simulation", fill=(0, 51, 102), font=small)
    d.line((420, 270, 500, 270), fill=(0, 51, 102), width=8)
    d.polygon([(500, 250), (540, 270), (500, 290)], fill=(0, 51, 102))
    d.rounded_rectangle((560, 80, 860, 460), radius=16, outline=(0, 51, 102), width=4)
    d.text((600, 180), "Reality", fill=(0, 51, 102), font=font)
    d.text((600, 230), "on the robot", fill=(0, 51, 102), font=small)
    im.save(SIM / "offer_bridge.png")

    # Physics chain.
    im = canvas(520, 360, (48, 48, 52))
    d = ImageDraw.Draw(im)
    for i, x in enumerate(range(40, 400, 46)):
        d.ellipse((x, 150, x + 36, 186), fill=(210, 170, 120))
    d.rectangle((430, 120, 490, 170), fill=(220, 120, 140))
    d.rectangle((430, 200, 490, 250), fill=(220, 120, 140))
    im.save(SIM / "physics_chain.png")

    # Renderer.
    im = canvas(520, 360, (28, 28, 32))
    d = ImageDraw.Draw(im)
    d.polygon([(40, 300), (480, 250), (480, 330), (40, 340)], fill=(90, 70, 50))
    d.rectangle((180, 80, 230, 220), fill=(180, 180, 186))
    d.polygon([(230, 150), (340, 190), (320, 220), (210, 180)], fill=(200, 80, 70))
    im.save(SIM / "renderer_view.png")

    # Sensors.
    im = canvas(520, 360, (70, 74, 80))
    d = ImageDraw.Draw(im)
    for i, (x, y) in enumerate(((30, 30), (270, 30), (30, 190), (270, 190))):
        d.rectangle((x, y, x + 210, y + 140), fill=(90, 40, 130))
        d.rectangle((x + 40, y + 40, x + 160, y + 100), fill=(200, 80, 50))
    im.save(SIM / "sensors_view.png")

    # UI.
    im = canvas(520, 360, (32, 36, 42))
    d = ImageDraw.Draw(im)
    d.rectangle((16, 16, 150, 344), fill=(48, 32, 40))
    for y in range(40, 300, 28):
        d.rectangle((28, y, 136, y + 16), fill=(120, 70, 80))
    d.rectangle((166, 16, 504, 344), fill=(70, 120, 160))
    d.rectangle((250, 40, 290, 250), fill=(210, 150, 80))
    im.save(SIM / "ui_panel.png")

    # A rendered-looking arm on a grid.
    im = canvas(860, 520, (214, 226, 232))
    d = ImageDraw.Draw(im)
    for i in range(0, 860, 40):
        d.line((i, 300, i - 80, 500), fill=(170, 180, 186), width=2)
    for i in range(220, 520, 28):
        d.line((0, i, 860, i - 40), fill=(170, 180, 186), width=1)
    d.rectangle((80, 250, 780, 300), fill=(230, 190, 190))
    d.polygon([(430, 80), (470, 80), (500, 220), (400, 220)], fill=(200, 202, 206))
    d.polygon([(500, 210), (620, 250), (600, 280), (470, 240)], fill=(230, 190, 40))
    d.ellipse((600, 240, 650, 290), fill=(40, 40, 44))
    im.save(SIM / "render_scene.png")

    im = canvas(860, 480, (20, 24, 32))
    d = ImageDraw.Draw(im)
    for i in range(8):
        shade = 30 + i * 18
        d.rectangle((0, 360 - i * 28, 860, 392 - i * 28), fill=(shade, shade + 20, shade + 40))
    d.polygon([(80, 300), (420, 80), (470, 110), (160, 320)], fill=(210, 212, 216))
    d.rectangle((150, 300, 230, 360), fill=(180, 60, 50))
    d.rectangle((520, 200, 760, 340), fill=(150, 150, 156))
    im.save(SIM / "mujoco_arm.png")

    im = canvas(640, 360, (230, 230, 228))
    d = ImageDraw.Draw(im)
    for i in range(12):
        d.line((20, 40 + i * 26, 300, 20 + i * 26), fill=(80, 60, 40), width=8)
    import math

    d.ellipse((360, 140, 560, 250), fill=(190, 140, 90))
    for ang in range(0, 360, 45):
        rad = math.radians(ang)
        d.line((460, 195, 460 + int(90 * math.cos(rad)), 195 + int(48 * math.sin(rad))), fill=(160, 110, 70), width=10)
    im.save(SIM / "mujoco_soft.png")

    im = canvas(640, 360, (20, 36, 56))
    d = ImageDraw.Draw(im)
    for i in range(8):
        for j in range(8):
            d.ellipse((180 + i * 36, 40 + j * 36, 208 + i * 36, 68 + j * 36), fill=(190, 40, 40))
    d.ellipse((80, 250, 150, 320), fill=(220, 220, 220))
    im.save(SIM / "mujoco_particles.png")

    # Benchmark schematic. Values follow the shape of the lecture figure.
    im = canvas(980, 620, (255, 255, 255))
    d = ImageDraw.Draw(im)
    left, top, right, bottom = 90, 70, 930, 540
    d.line((left, top, left, bottom), fill=(80, 80, 80), width=2)
    d.line((left, bottom, right, bottom), fill=(80, 80, 80), width=2)
    for frac, label in ((0, "0"), (0.25, "200k"), (0.5, "400k"), (0.75, "600k"), (1, "800k")):
        y = bottom - int((bottom - top) * frac)
        d.line((left, y, right, y), fill=(220, 220, 220), width=1)
        d.text((8, y - 12), label, fill=(80, 80, 80), font=small)
    mjx = [(1, 0.82), (2, 0.35), (3, 0.18), (4, 0.12), (6, 0.06), (8, 0.04), (10, 0.03)]
    mj = [(1, 0.18), (2, 0.09), (4, 0.05), (6, 0.035), (8, 0.028), (10, 0.02)]

    def xy(points):
        out = []
        for n, v in points:
            x = left + int((right - left) * (n / 10))
            y = bottom - int((bottom - top) * v)
            out.append((x, y))
        return out

    d.line(xy(mj), fill=(40, 90, 180), width=4)
    d.line(xy(mjx), fill=(190, 40, 40), width=4)
    d.text((300, 16), "MuJoCo vs MJX, large scene, steps/second", fill=(90, 90, 90), font=small)
    d.rectangle((620, 48, 648, 64), fill=(40, 90, 180))
    d.text((656, 42), "MuJoCo", fill=(40, 40, 40), font=small)
    d.rectangle((780, 48, 808, 64), fill=(190, 40, 40))
    d.text((816, 42), "MJX", fill=(40, 40, 40), font=small)
    d.text((430, 560), "# Humanoids", fill=(60, 60, 60), font=small)
    im.save(SIM / "mjx_chart.png")

    # Six classic Gym tasks, drawn rather than copied from the docs.
    im = canvas(980, 640, (255, 255, 255))
    d = ImageDraw.Draw(im)
    labels = ["CartPole-v1", "MountainCar-v0", "Pendulum-v1", "Reacher-v4", "InvertedPendulum-v4", "Acrobot-v1"]
    for i, label in enumerate(labels):
        c, r = i % 3, i // 3
        x, y = 20 + c * 320, 16 + r * 310
        d.rounded_rectangle((x, y, x + 300, y + 290), radius=12, outline=(180, 180, 180), width=2)
        d.rectangle((x + 20, y + 30, x + 280, y + 210), fill=(245, 245, 245))
        d.text((x + 16, y + 230), label, fill=(30, 30, 30), font=small)
    # tiny marks so each panel is distinguishable
    d.rectangle((70, 150, 90, 200), fill=(40, 40, 40))
    d.line((80, 150, 110, 90), fill=(40, 40, 40), width=6)
    d.polygon([(360, 180), (520, 80), (560, 120)], outline=(40, 40, 40))
    d.ellipse((760, 80, 860, 180), outline=(180, 40, 40), width=6)
    d.line((1040, 140, 1120, 180), fill=(40, 120, 80), width=6)
    d.rectangle((1400, 160, 1420, 220), fill=(40, 40, 40))
    d.line((1680, 80, 1740, 160), fill=(40, 40, 40), width=6)
    d.line((1740, 160, 1800, 90), fill=(40, 40, 40), width=6)
    im.save(SIM / "gym_tasks.png")


def _why(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Why simulation?")
    slide.shapes.add_movie(
        str(SIM / "atlas_fail.mp4"),
        _emu(0.38),
        _emu(0.92),
        _emu(4.48),
        _emu(2.55),
        poster_frame_image=str(SIM / "atlas_fail.jpg"),
        mime_type="video/mp4",
    )
    _box(
        slide,
        0.38,
        3.54,
        4.48,
        0.78,
        ["Example 1: testing a controller on an expensive robot without knowing if it will crash"],
        size=16,
    )
    _box(
        slide,
        0.38,
        4.32,
        4.48,
        0.32,
        ["Boston Dynamics, Farewell to HD Atlas, YouTube, 2024."],
        size=11,
        color=MUTED,
    )
    slide.shapes.add_picture(str(SIM / "dactyl_hand.jpg"), _emu(5.12), _emu(0.92), _emu(4.48), _emu(2.55))
    _box(
        slide,
        5.12,
        3.54,
        4.48,
        0.78,
        ["Example 2: an RL policy that needed about 13,000 years of experience in simulation"],
        size=16,
    )
    _box(
        slide,
        5.12,
        4.32,
        4.48,
        0.32,
        ["OpenAI Dactyl, Rubik's Cube, 2019. Photo: Shadow Hand, CC BY-SA 3.0."],
        size=11,
        color=MUTED,
    )
    _credit(slide, "Framework after " + CREDIT)


def _offers(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "What simulation offers")
    cards = [
        (SIM / "nv_warehouse.jpg", "Bridging design and reality", 0.38, 0.92),
        (SIM / "nv_franka.jpg", "Safe and repeatable testing", 5.12, 0.92),
        (SIM / "nv_scale.jpg", "Scalability", 0.38, 2.95),
        (SIM / "nv_truth.jpg", "Access to ground truth signals", 5.12, 2.95),
    ]
    for path, caption, x, y in cards:
        slide.shapes.add_picture(str(path), _emu(x), _emu(y), _emu(4.48), _emu(1.55))
        _box(slide, x, y + 1.55, 4.48, 0.36, [("•  " + caption, False, 16, INK)], size=16)
    _box(
        slide,
        0.38,
        4.78,
        9.2,
        0.28,
        ["Images: NVIDIA Isaac Sim, Isaac Gym, and Factory. NVIDIA Technical Blog."],
        size=11,
        color=MUTED,
    )
    _credit(slide)


def _practice(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Robotic simulation in practice")
    items = [
        ("Classical control verification", "Study the performance, stability, and robustness of control systems"),
        ("Data generation", "Generating large datasets for robot policy training"),
        ("Evaluation", "Benchmark algorithms under controlled, repeatable conditions"),
        ("Design and prototyping", "Rapidly test robot morphology, sensor placement, or gripper designs"),
    ]
    y = 1.05
    for head, body in items:
        _box(
            slide,
            0.55,
            y,
            8.8,
            0.78,
            [
                [("•  " + head, True, 20, INK)],
                [("     " + body, False, 18, INK)],
            ],
        )
        y += 0.9
    _credit(slide)


def _card(slide, x, y, w, h, heading, image, bullets) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x), _emu(y), _emu(w), _emu(h))
    shape.adjustments[0] = 0.08
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(0x22, 0x22, 0x22)
    shape.line.width = Pt(1.5)
    _box(slide, x + 0.15, y + 0.08, w - 0.3, 0.38, [(heading, True, 18, INK)], align=PP_ALIGN.CENTER)
    slide.shapes.add_picture(str(image), _emu(x + 0.18), _emu(y + 0.52), _emu(1.7), _emu(1.18))
    lines = [[("•  " + b, False, 14, INK)] for b in bullets]
    _box(slide, x + 2.0, y + 0.52, w - 2.2, h - 0.65, lines, size=14)


def _components(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Components of a simulator")
    _card(slide, 0.32, 0.95, 4.6, 1.9, "Physics Engine", SIM / "comp_physics.jpg", ["Rigid, articulated, deformable, …", "Speed: CPU/GPU", "Accuracy"])
    _card(slide, 5.08, 0.95, 4.6, 1.9, "Renderer", SIM / "comp_renderer.jpg", ["Method: rasterization, path tracing, …", "Lighting, texture, …", "Speed"])
    _card(slide, 0.32, 2.98, 4.6, 1.9, "Sensor Simulation", SIM / "comp_sensors.jpg", ["RGB camera", "Depth/LiDAR", "Robot states", "Tactile"])
    _card(slide, 5.08, 2.98, 4.6, 1.9, "User Interface", SIM / "comp_ui.jpg", ["Scene setup", "API", "Policy training"])
    _box(
        slide,
        0.38,
        4.95,
        8.3,
        0.22,
        ["Images: MuJoCo documentation (Google DeepMind); robosuite documentation."],
        size=10,
        color=MUTED,
    )
    _credit(slide)


def _labeled_box(slide, x, y, w, h, lines, fill) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x), _emu(y), _emu(w), _emu(h))
    shape.adjustments[0] = 0.1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = RGBColor(0x88, 0x88, 0x88)
    shape.line.width = Pt(1)
    tf = shape.text_frame
    tf.word_wrap = True
    tf._txBody.find(qn("a:bodyPr")).set("anchor", "ctr")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        add_muli_run(p, line, "Muli", 11, False, INK)


def _step(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "The step() function")
    _box(
        slide,
        0.48,
        0.88,
        9.0,
        1.35,
        [
            [("•  Update the simulation state from  t  to  t + Δt", False, 16, INK)],
            [("•  Physics step:", False, 16, INK)],
            [("      •  Detecting collision", False, 16, INK)],
            [("      •  Solving physics (resolve collision, evolve dynamics)", False, 16, INK)],
            [("      •  Time integration", False, 16, INK)],
        ],
    )
    slide.shapes.add_picture(str(SIM / "nerd_workflow.png"), _emu(0.70), _emu(2.22), _emu(8.60), _emu(2.60))
    _box(
        slide,
        0.38,
        4.86,
        8.4,
        0.24,
        ["Xu, J., Heiden, E., Akinola, I., Fox, D., Macklin, M., and Narang, Y. Neural Robot Dynamics. arXiv:2508.15755, 2025. Fig. 2(a)."],
        size=10,
        color=MUTED,
    )
    _credit(slide)


def _render(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "The render() function")
    _box(
        slide,
        0.48,
        1.05,
        4.6,
        2.6,
        [
            [("•  Get rendering and other sensor readings", False, 18, INK)],
            [("•  Rendering step:", False, 18, INK)],
            [("      •  Sending simulation states to the renderer", False, 18, INK)],
            [("      •  Run rendering algorithm", False, 18, INK)],
            [("      •  Update viewer (if interactive)", False, 18, INK)],
        ],
    )
    slide.shapes.add_picture(str(SIM / "render_handover.jpg"), _emu(5.2), _emu(1.15), _emu(4.4), _emu(2.7))
    _box(
        slide,
        5.2,
        3.90,
        4.4,
        0.32,
        ["robosuite TwoArmHandover. robosuite documentation."],
        size=10,
        color=MUTED,
        align=PP_ALIGN.CENTER,
    )
    _credit(slide)


def _contain_picture(slide, path, x, y, w, h) -> None:
    iw, ih = Image.open(path).size
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    slide.shapes.add_picture(str(path), _emu(x + (w - pw) / 2), _emu(y + (h - ph) / 2), _emu(pw), _emu(ph))


def _simulators(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Major robotic simulators")
    logos = [
        (SIM / "logos" / "mujoco.png", 0.40, 0.95),
        (SIM / "logos" / "isaac.png", 5.15, 0.95),
        (SIM / "logos" / "sapien.png", 0.40, 2.90),
        (SIM / "logos" / "genesis.png", 5.15, 2.90),
    ]
    for path, x, y in logos:
        _contain_picture(slide, path, x, y, 4.40, 1.70)
    _box(
        slide,
        0.38,
        4.78,
        8.6,
        0.28,
        ["Logos: MuJoCo (Google DeepMind); NVIDIA Isaac Lab; SAPIEN (UCSD); Genesis."],
        size=10,
        color=MUTED,
    )
    _credit(slide)


def _mujoco(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "MuJoCo")
    _box(
        slide,
        0.42,
        0.88,
        4.7,
        3.95,
        [
            [("•  Multi-Joint dynamics with Contact", True, 17, INK)],
            [("•  A physics engine for robotics, biomechanics, and machine learning", False, 15, INK)],
            [("•  Written at the University of Washington, 2012", False, 15, INK)],
            [("      Todorov, Erez, and Tassa, IROS 2012", False, 14, MUTED)],
            [("•  Sold as Roboti LLC, then acquired by DeepMind in 2021", False, 15, INK)],
            [("•  Open-sourced in 2022 under Apache 2.0", False, 15, INK)],
            [("•  Physics engine: MuJoCo", False, 15, INK)],
            [("•  Rendering: rasterization", False, 15, INK)],
        ],
    )
    _box(
        slide,
        5.15,
        0.88,
        4.4,
        3.95,
        [
            [("•  Strengths", True, 17, INK)],
            [("      •  Fast CPU physics", False, 15, INK)],
            [("      •  Smooth contact", False, 15, INK)],
            [("      •  Light interface", False, 15, INK)],
            [("•  Weaknesses", True, 17, INK)],
            [("      •  Limited rendering", False, 15, INK)],
            [("      •  Single thread", False, 15, INK)],
            [("•  MJX / MJWarp", True, 17, INK)],
            [("      •  Multi-thread, GPU acceleration", False, 15, INK)],
        ],
    )
    _box(
        slide,
        0.42,
        4.86,
        8.4,
        0.24,
        ["Wikipedia, “MuJoCo.” Todorov, Erez, and Tassa, IROS 2012."],
        size=10,
        color=MUTED,
    )
    _credit(slide)


def _tools(prs, exemplar: int) -> None:
    slide = _fresh(prs, exemplar)
    _title(slide, "Other tools")
    _box(slide, 0.42, 0.88, 2.20, 0.50, [[("•  OpenAI Gym  /", True, 18, INK)]], anchor="ctr")
    _contain_picture(slide, SIM / "logos" / "gymnasium.png", 2.55, 0.88, 2.35, 0.50)
    _box(
        slide,
        0.42,
        1.42,
        4.55,
        3.15,
        [
            [("•  An API standard for reinforcement learning", False, 16, INK)],
            [("•  Core: Step function", True, 18, INK)],
            [("      •  Policy rollout", False, 16, INK)],
            [("      •  Env stepping", False, 16, INK)],
            [("      •  Obtain observation and reward", False, 16, INK)],
            [("•  Also: Render, Make, Reset", True, 18, INK)],
        ],
    )
    slide.shapes.add_picture(str(SIM / "gym_envs.jpg"), _emu(5.10), _emu(0.86), _emu(4.50), _emu(1.68))
    _contain_picture(slide, SIM / "gym_loop.png", 5.10, 2.60, 4.50, 2.30)
    _box(
        slide,
        5.10,
        4.94,
        4.50,
        0.20,
        ["Gymnasium, Farama Foundation. Environment stills from Gymnasium documentation."],
        size=9,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )
    _credit(slide)


def add_simulation_slides(prs, exemplar: int) -> None:
    _make_images()
    _section(prs, exemplar)
    _why(prs, exemplar)
    _offers(prs, exemplar)
    _practice(prs, exemplar)
    _components(prs, exemplar)
    _step(prs, exemplar)
    _render(prs, exemplar)
    _simulators(prs, exemplar)
    _mujoco(prs, exemplar)
    _tools(prs, exemplar)
