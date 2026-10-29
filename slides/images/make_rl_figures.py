#!/usr/bin/env python3
"""Simple navy diagrams for the two notation / reward slides."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent
NAVY = (0, 51, 102)
GOLD = (196, 163, 90)
INK = (38, 38, 38)
MUTED = (107, 107, 107)
PALE = (232, 238, 244)
WHITE = (255, 255, 255)


def _font(size: int, bold: bool = False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _center(draw, xy, text, font, fill):
    box = draw.textbbox((0, 0), text, font=font)
    w, h = box[2] - box[0], box[3] - box[1]
    draw.text((xy[0] - w / 2, xy[1] - h / 2), text, font=font, fill=fill)


def _arrow(draw, p0, p1, color=NAVY, width=3):
    draw.line([p0, p1], fill=color, width=width)
    x0, y0 = p0
    x1, y1 = p1
    import math

    ang = math.atan2(y1 - y0, x1 - x0)
    L = 12
    left = (x1 - L * math.cos(ang - 0.45), y1 - L * math.sin(ang - 0.45))
    right = (x1 - L * math.cos(ang + 0.45), y1 - L * math.sin(ang + 0.45))
    draw.polygon([p1, left, right], fill=color)


def _circle(draw, xy, r, fill, outline, width=3):
    x, y = xy
    draw.ellipse((x - r, y - r, x + r, y + r), fill=fill, outline=outline, width=width)


def make_loop():
    W, H = 1440, 900
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)
    label_f = _font(28, True)
    small_f = _font(24, False)
    tiny_f = _font(22, False)

    # One column per time step: o --π--> a on top, s below.
    xs = [220, 720, 1220]
    y_top, y_s = 220, 680
    r = 58
    dx = 150

    for i, x in enumerate(xs, 1):
        xo, xa = x - dx, x + dx
        _circle(d, (xo, y_top), r, PALE, NAVY)
        _center(d, (xo, y_top), f"o{i}", label_f, NAVY)
        _circle(d, (xa, y_top), r, WHITE, NAVY)
        _center(d, (xa, y_top), f"a{i}", label_f, NAVY)
        _circle(d, (x, y_s), r, NAVY, NAVY)
        _center(d, (x, y_s), f"s{i}", label_f, WHITE)
        _arrow(d, (x, y_s - r - 6), (xo, y_top + r + 8), color=MUTED, width=2)
        _arrow(d, (xo + r + 8, y_top), (xa - r - 8, y_top))
        _center(d, ((xo + xa) / 2, y_top - 42), "π", small_f, GOLD)

    for i in range(2):
        x0, x1 = xs[i], xs[i + 1]
        _arrow(d, (x0 + r + 8, y_s), (x1 - r - 8, y_s))
        _center(d, ((x0 + x1) / 2, y_s + 48), "p(s'|s, a)", tiny_f, MUTED)

    im.save(OUT / "rl_loop_notation.png")


def make_reward():
    W, H = 1440, 1080
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)
    title_f = _font(32, True)
    body_f = _font(24, False)
    small_f = _font(22, False)

    def card(box, title, lines, edge):
        d.rounded_rectangle(box, radius=22, fill=WHITE, outline=edge, width=4)
        x0, y0, x1, y1 = box
        d.text((x0 + 28, y0 + 22), title, font=title_f, fill=edge)
        for i, line in enumerate(lines):
            d.text((x0 + 28, y0 + 80 + 38 * i), line, font=body_f, fill=INK)

    card((40, 40, 690, 430), "High reward", ["Tesla: stay in lane, miss the truck", "AlphaGo: the game is won"], NAVY)
    card((750, 40, 1400, 430), "Low reward", ["Tesla: hit the truck, leave the road", "AlphaGo: the game is lost"], GOLD)

    d.rounded_rectangle((40, 480, 1400, 1040), radius=22, fill=PALE, outline=None)
    d.text((80, 530), "r(s, a)  says which (s, a) is better", font=title_f, fill=NAVY)
    d.text((80, 640), "An MDP is  (S,  A,  p,  r,  γ)", font=title_f, fill=INK)
    d.text((80, 740), "state, action, next-state rule, reward, discount", font=small_f, fill=MUTED)
    d.text((80, 860), "Return = add the rewards along the episode", font=body_f, fill=INK)
    im.save(OUT / "reward_mdp.png")


if __name__ == "__main__":
    make_loop()
    make_reward()
    print("wrote", OUT / "rl_loop_notation.png")
    print("wrote", OUT / "reward_mdp.png")
