"""Validated, offline badge drawing shared by preview and downloaded cards.

Coordinates are always on a 600 x 900 canvas. Text positions are centre-line
anchors; image positions are top-left corners. No network or event writes here.
"""
import math
import os
import re
from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageOps

WIDTH, HEIGHT = 600, 900


def resolve_design(config, row, groups=()):
    """Named groups take priority over subgroup designs, then the event default.

    Support both the current designer stores and legacy communication templates.
    """
    normalized = {str(k).strip().lower(): str(v or "").strip() for k, v in row.items()}
    registration = normalized.get("registration number", "")
    for group in groups:
        if registration and registration in group.get("reg_nos", []):
            design = config.get("custom_group_designer_settings", {}).get(group.get("name")) or group.get("id_card_designer_settings")
            if design:
                return design
    column = config.get("group_column", "").strip()
    value = normalized.get(column.lower())
    if column and value:
        key = f"{column}:{value}"
        design = config.get("subgroup_designer_settings", {}).get(key) or config.get("subgroup_templates", {}).get(key, {}).get("id_card_designer_settings")
        if design:
            return design
    return config.get("id_card_designer_settings") or {}


def color(value, default=(255, 255, 255, 255)):
    try:
        if isinstance(value, str):
            rgba = re.fullmatch(r"rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)", value)
            if rgba:
                r, g, b, a = rgba.groups()
                value = (int(r), int(g), int(b), round(float(a) * 255))
            else:
                value = ImageColor.getcolor(value, "RGBA")
        values = tuple(max(0, min(255, int(v))) for v in value)
        if len(values) == 3:
            return values + (255,)
        if len(values) == 4:
            return values
    except (TypeError, ValueError):
        pass
    return default


def number(value, default, low, high):
    try:
        value = float(value)
        if math.isfinite(value):
            return max(low, min(high, round(value)))
    except (TypeError, ValueError, OverflowError):
        pass
    return default


def validate_design(settings):
    if not isinstance(settings, dict):
        raise ValueError("Choose a valid card design.")
    if settings.get("mode", "custom") not in ("custom", "custom_html"):
        raise ValueError("Unsupported card design mode.")
    fields = settings.get("fields", [])
    if not isinstance(fields, list) or len(fields) > 60 or any(not isinstance(f, dict) for f in fields):
        raise ValueError("A card must contain a list of at most 60 fields.")
    for f in fields:
        if f.get("type", "text") not in ("text", "image"):
            raise ValueError("Card fields must be text or images.")
        if len(str(f.get("text", ""))) > 2000:
            raise ValueError("Keep each card text field under 2,000 characters.")
    for key in ("qr", "photo"):
        if key in settings and not isinstance(settings[key], dict):
            raise ValueError(f"Invalid {key} settings.")
    shapes = settings.get("shapes", [])
    if not isinstance(shapes, list) or len(shapes) > 60 or any(not isinstance(s, dict) for s in shapes):
        raise ValueError("A card must contain a list of at most 60 shapes.")
    return settings


def field_text(field, context):
    def key(value):
        return re.sub(r"[\s_]+", " ", str(value or "").strip().lower())
    binding = field.get("bind_column")
    if binding:
        # An absent value must never print a sample label (e.g. VIP PASS).
        return str(context.get(key(binding), "") or "")
    return re.sub(r"\{([^{}]{1,80})\}", lambda m: str(context.get(key(m[1]), "") or ""), str(field.get("text", "")))


def font(size, family="Arial", bold=False, italic=False):
    mono = any(s in family.lower() for s in ("mono", "courier", "consolas"))
    serif = any(s in family.lower() for s in ("georgia", "times", "playfair"))
    base = "cour" if mono else "georgia" if serif else "arial"
    suffix = ("bi" if bold and italic else "bd" if bold else "i" if italic else "")
    if serif:
        suffix = "z" if bold and italic else "b" if bold else "i" if italic else ""
    dejavu = "DejaVuSansMono" if mono else "DejaVuSerif" if serif else "DejaVuSans"
    variant = "-BoldOblique" if bold and italic else "-Bold" if bold else "-Oblique" if italic else ""
    for path in (os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", base + suffix + ".ttf"),
                 dejavu + variant + ".ttf", "/usr/share/fonts/truetype/dejavu/" + dejavu + variant + ".ttf"):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def render_custom(settings, context, qr_path, image_loader=None, photo=None):
    validate_design(settings)
    start = color(settings.get("bg_start"), (11, 15, 25, 255))
    end = color(settings.get("bg_end"), (26, 35, 51, 255))
    card = Image.new("RGBA", (WIDTH, HEIGHT))
    draw = ImageDraw.Draw(card)
    for y in range(HEIGHT):
        t = y / (HEIGHT - 1)
        draw.line((0, y, WIDTH, y), fill=tuple(round(start[i] * (1-t) + end[i] * t) for i in range(3)) + (255,))
    layer = Image.new("RGBA", card.size)
    detail = ImageDraw.Draw(layer)
    for shape in settings.get("shapes", []):
        fill = color(shape.get("color"), (255, 255, 255, 20))
        box = shape.get("box")
        stroke = number(shape.get("width"), 2, 1, 100)
        try:
            kind = shape.get("type")
            if kind == "arc":
                detail.arc(box, 0, 360, fill=fill, width=stroke)
            elif kind == "circle":
                detail.ellipse(box, fill=fill)
            elif kind == "rectangle":
                detail.rectangle(box, fill=fill)
            elif kind == "line":
                detail.line([tuple(p) for p in shape.get("coords", [])], fill=fill, width=stroke)
        except (TypeError, ValueError):
            continue
    detail.rounded_rectangle((16, 16, 584, 884), radius=24,
                             outline=color(settings.get("border_color"), (255, 255, 255, 30)),
                             width=number(settings.get("border_width"), 2, 1, 20))
    card = Image.alpha_composite(card, layer)
    draw = ImageDraw.Draw(card)
    accent = color(settings.get("accent_color"), (0, 245, 160, 255))
    if settings.get("photo", {}).get("enabled") and photo is not False:
        spec = settings["photo"]
        size = number(spec.get("size"), 150, 40, 300)
        x = number(spec.get("x"), 225, 24, WIDTH-size-24)
        y = number(spec.get("y"), 135, 24, HEIGHT-size-24)
        avatar = Image.new("RGB", (size, size), accent[:3])
        if photo is not None:
            avatar = ImageOps.fit(photo.convert("RGB"), (size, size))
        else:
            initials = "".join(w[0] for w in context.get("name", "Guest").split()[:2]).upper()
            ImageDraw.Draw(avatar).text((size/2, size/2), initials, font=font(size//3, bold=True), fill=start[:3], anchor="mm")
        mask = Image.new("L", (size, size))
        ImageDraw.Draw(mask).ellipse((0, 0, size-1, size-1), fill=255)
        card.paste(avatar, (x, y), mask)
    for field in settings.get("fields", []):
        x = number(field.get("x"), 300, 24, WIDTH-24)
        y = number(field.get("y"), 210, 24, HEIGHT-24)
        if field.get("type") == "image":
            source = image_loader(field) if image_loader else None
            if source is not None:
                w = number(field.get("width"), 100, 1, WIDTH-x)
                h = number(field.get("height"), 100, 1, HEIGHT-y)
                overlay = ImageOps.contain(source.convert("RGBA"), (w, h))
                card.paste(overlay, (x, y), overlay)
            continue
        text = field_text(field, context).replace("\n", " ")
        transform = field.get("transform")
        text = text.upper() if transform == "uppercase" else text.lower() if transform == "lowercase" else text.title() if transform == "capitalize" else text
        size = number(field.get("font_size"), 20, 8, 120)
        align = field.get("align", "center")
        limit = (WIDTH-x-24) if align == "left" else (x-24) if align == "right" else 2 * min(x-24, WIDTH-x-24)
        limit = max(1, limit)
        def get_font(s):
            return font(s, str(field.get("font_name", "Arial")), field.get("bold", False), field.get("italic", False))
        spacing_ratio = {'wide': 0.1, 'extra_wide': 0.25}.get(field.get('tracking'), 0)
        def text_width(value):
            if not spacing_ratio:
                return draw.textlength(value, font=face)
            return sum(draw.textlength(char, font=face) for char in value) + max(0, len(value)-1) * size * spacing_ratio
        face = get_font(size)
        while size > 10 and text_width(text) > limit:
            size -= 1
            face = get_font(size)
        if text_width(text) > limit:
            while text and text_width(text + "…") > limit:
                text = text[:-1]
            text += "…"
        anchor = "lm" if align == "left" else "rm" if align == "right" else "mm"
        fill = color(field.get("color"))
        measured = text_width(text)
        left = x if align == 'left' else x-measured if align == 'right' else x-measured/2
        if spacing_ratio:
            cursor = left
            for char in text:
                draw.text((cursor, y), char, font=face, fill=fill, anchor='lm')
                cursor += draw.textlength(char, font=face) + size * spacing_ratio
        else:
            draw.text((x, y), text, font=face, fill=fill, anchor=anchor)
        if field.get("underline") and text:
            bounds = draw.textbbox((x, y), text, font=face, anchor=anchor)
            draw.line((left, bounds[3]+2, left+measured, bounds[3]+2), fill=fill, width=max(1, size//12))
    # Draw QR last, on opaque white with its original quiet zone intact.
    qr = settings.get("qr", {})
    size = number(qr.get("size"), 240, 160, 360)
    x = number(qr.get("x"), 180, 24, WIDTH-size-24)
    y = number(qr.get("y"), 550, 24, HEIGHT-size-24)
    if qr_path:
        with Image.open(qr_path) as source:
            code = source.convert("RGB").resize((size, size), Image.Resampling.NEAREST)
        draw.rounded_rectangle((x-10, y-10, x+size+10, y+size+10), radius=14, fill="white")
        card.paste(code, (x, y))
    return card.convert("RGB")
