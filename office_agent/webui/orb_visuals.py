from functools import lru_cache
import math
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont


STATUS_LABELS = {'ready': '拖入文档', 'opening': '正在打开', 'busy': '处理中', 'complete': '副本已验证',
                 'confirm': '等待你确认', 'uncertain': '需要核对结果',
                 'error': '操作未完成', 'stopped': '已停止后续步骤'}
ACCENTS = {'opening': (220, 220, 220), 'busy': (220, 220, 220),
           'confirm': (245, 185, 66), 'uncertain': (255, 170, 163), 'error': (255, 170, 163)}


def display_status(stage, busy, opening, confirmation):
    if stage == 'uncertain':
        return 'uncertain'
    if stage == 'stopped':
        return 'stopped'
    if stage in {'failed', 'open_failed'}:
        return 'error'
    return 'opening' if opening else 'busy' if busy else 'confirm' if confirmation else 'complete' if stage == 'verified' else 'ready'


@lru_cache(maxsize=4)
def points(count=192):
    result = []
    for index in range(count):
        y = 1 - 2 * (index + .5) / count
        radius = math.sqrt(max(0, 1 - y * y))
        angle = index * math.pi * (3 - math.sqrt(5))
        result.append((math.cos(angle) * radius, y, math.sin(angle) * radius))
    return tuple(result)


def action_positions(origin, scale, area):
    size = round(40 * scale)
    center = (origin[0] + 32 * scale, origin[1] + 32 * scale)
    left, top, right, bottom = area
    inward = math.atan2((top + bottom) / 2 - center[1], (left + right) / 2 - center[0])
    directions = [inward] + [math.radians(a) for a in range(-180, 180, 15)]
    for span in (120, 110, 90):
        for direction in directions:
            angles = [direction + math.radians(-span / 2 + span * i / 3) for i in range(4)]
            positions = [(round(center[0] + 96 * scale * math.cos(a) - size / 2),
                          round(center[1] + 96 * scale * math.sin(a) - size / 2)) for a in angles]
            if all(left <= x and top <= y and x + size <= right and y + size <= bottom for x, y in positions):
                return positions
    raise ValueError('屏幕可用区域不足，请移动悬浮球后展开快捷操作')


@lru_cache(maxsize=128)
def orb_image(size=64, status='ready', phase=0, active=False, focused=False, highcontrast=False, pulse=0, emphasis=0):
    if not 48 <= size <= 256 or status not in STATUS_LABELS:
        raise ValueError('Invalid orb appearance')
    factor = 3
    canvas = Image.new('RGBA', (size * factor, size * factor))
    draw = ImageDraw.Draw(canvas)
    center = size * factor / 2
    radius = size * factor * (.415 + .0015 * pulse + .0018 * emphasis)
    color = ACCENTS.get(status, (237, 241, 248))
    if active:
        color = (240, 240, 240)
    draw.ellipse((center - radius, center - radius, center + radius, center + radius),
                 fill=(18, 18, 18, 210 if highcontrast else 145))
    angle = phase * math.tau / 360
    transformed = [(x * math.cos(angle) + z * math.sin(angle), y,
                    z * math.cos(angle) - x * math.sin(angle)) for x, y, z in points()]
    for x, y, z in sorted(transformed, key=lambda p: p[2]):
        light = (z + 1) / 2
        perspective = .83 + .17 * light
        diameter = (.65 + 1.65 * light ** 1.6) * size / 64 * factor
        px = center + x * radius * perspective
        py = center + y * radius * perspective
        alpha = min(255, round(32 + (213 + emphasis * 1.25) * light ** 1.35))
        draw.ellipse((px - diameter / 2, py - diameter / 2, px + diameter / 2, py + diameter / 2), fill=(*color, alpha))
    if active or focused or highcontrast or status in {'confirm', 'uncertain', 'error'}:
        draw.ellipse((center - radius - factor, center - radius - factor,
                      center + radius + factor, center + radius + factor),
                     outline=(*color, 240 if focused or highcontrast else 155), width=factor)
    result = canvas.resize((size, size), Image.Resampling.LANCZOS)
    mask = Image.new('L', (size, size))
    ImageDraw.Draw(mask).ellipse((1, 1, size - 2, size - 2), fill=255)
    result.putalpha(ImageChops.multiply(result.getchannel('A'), mask))
    return result


def decoration_image(size, status, visible_label=False, active=False, anchor=None):
    width, height = round(size * 2), round(size * 1.95)
    result = Image.new('RGBA', (width, height))
    color = ACCENTS.get(status, (190, 190, 190))
    if active:
        color = (220, 220, 220)
    glow = Image.new('RGBA', (width, height))
    draw = ImageDraw.Draw(glow)
    center_x, center_y = anchor or (width / 2, size * .72)
    radius = size * .45
    opacity = 34 if status == 'ready' and not active else 62
    draw.ellipse((center_x - radius, center_y - radius, center_x + radius, center_y + radius), fill=(*color, opacity))
    result.alpha_composite(glow.filter(ImageFilter.GaussianBlur(size * .2)))
    if visible_label or status != 'ready':
        label = STATUS_LABELS[status]
        if active:
            label = '松开即可打开'
        font_path = Path('C:/Windows/Fonts/msyh.ttc')
        font = ImageFont.truetype(str(font_path), max(10, round(size * .17)))
        draw = ImageDraw.Draw(result)
        box = draw.textbbox((0, 0), label, font=font)
        x = max(7, min(center_x - box[2] / 2, width - box[2] - 7))
        y = round(center_y + size * .67)
        if y + box[3] + 6 > height:
            y = max(2, round(center_y - size * .67 - box[3]))
        draw.rounded_rectangle((x - 6, y - 2, x + box[2] + 6, y + box[3] + 4), radius=5, fill=(20, 23, 30, 205))
        draw.text((x, y), label, font=font, fill=(237, 241, 248, 245))
    return result


@lru_cache(maxsize=32)
def action_image(size, icon, focused=False, highcontrast=False):
    if icon not in {'folder-open', 'message-square', 'history', 'settings', 'x'}:
        raise ValueError('Invalid action icon')
    factor = 3
    image = Image.new('RGBA', (size * factor, size * factor))
    draw = ImageDraw.Draw(image)
    draw.ellipse((factor, factor, (size - 1) * factor, (size - 1) * factor),
                 fill=(20, 20, 20, 255) if not focused else (42, 42, 42, 255),
                 outline=(245, 245, 245, 255) if focused or highcontrast else (110, 110, 110, 210),
                 width=factor if focused or highcontrast else max(1, factor // 2))
    image = image.resize((size, size), Image.Resampling.LANCZOS)
    asset = Path(__file__).resolve().parent / 'static/vendor/lucide/native' / (icon + '.png')
    with Image.open(asset) as source:
        glyph = source.convert('RGBA').resize((round(size * .48), round(size * .48)), Image.Resampling.LANCZOS)
    image.alpha_composite(glyph, ((size - glyph.width) // 2, (size - glyph.height) // 2))
    return image
