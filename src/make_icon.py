"""Иконка KompasRelink: два сцепленных звена цепи на сине-бирюзовом скруглённом квадрате."""
import os
from PIL import Image, ImageDraw, ImageFilter, ImageChops

S = 1024                      # рисуем крупно, потом уменьшаем
HERE = os.path.dirname(os.path.abspath(__file__))


def gradient(size, c1, c2):
    """Диагональный градиент c1 (левый верх) -> c2 (правый низ)."""
    g = Image.new("RGB", (size, size))
    px = g.load()
    for y in range(size):
        for x in range(size):
            t = (x + y) / (2 * (size - 1))
            px[x, y] = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
    return g


def rounded_mask(size, r):
    m = Image.new("L", (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, size - 1, size - 1), r, fill=255)
    return m


def link(draw, box, width, color):
    x0, y0, x1, y1 = box
    draw.rounded_rectangle(box, radius=(y1 - y0) // 2, outline=color, width=width)


def chain(size):
    """Слой с двумя звеньями (горизонтально), A поверх B снизу, B поверх A сверху."""
    L = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(L)
    W, sep = 92, 26                                   # толщина звена и «зазор» на пересечениях
    sep_col = (16, 70, 120, 255)
    A = (120, 362, 600, 662)                          # белое звено
    B = (424, 362, 904, 662)                          # оранжевое звено
    white, orange = (255, 255, 255, 255), (255, 152, 32, 255)

    def draw_link(dr, box, col):
        link(dr, (box[0] - sep // 2, box[1] - sep // 2, box[2] + sep // 2, box[3] + sep // 2), W + sep, sep_col)
        link(dr, box, W, col)

    draw_link(d, A, white)
    draw_link(d, B, orange)
    # нижнее пересечение: белое звено поверх оранжевого
    top = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw_link(ImageDraw.Draw(top), A, white)
    clip = Image.new("L", (size, size), 0)
    ImageDraw.Draw(clip).rectangle((380, 512, 660, size), fill=255)
    L.paste(top, (0, 0), ImageChops.multiply(clip, top.getchannel("A")))
    return L


def build():
    bg = gradient(S, (18, 92, 196), (0, 168, 170)).convert("RGBA")
    # мягкий блик сверху
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((-200, -620, S + 200, 430), fill=(255, 255, 255, 46))
    bg = Image.alpha_composite(bg, glow.filter(ImageFilter.GaussianBlur(40)))

    ch = chain(S).rotate(35, resample=Image.BICUBIC, center=(S / 2, S / 2))
    shadow = Image.new("RGBA", (S, S), (0, 20, 50, 0))
    shadow.putalpha(ch.getchannel("A").point(lambda a: int(a * 0.45)))
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    bg = Image.alpha_composite(bg, shadow.transform((S, S), Image.AFFINE, (1, 0, -10, 0, 1, -18)))
    bg = Image.alpha_composite(bg, ch)

    out = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    out.paste(bg, (0, 0), rounded_mask(S, 220))
    # тонкая светлая кромка
    edge = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(edge).rounded_rectangle((4, 4, S - 5, S - 5), 216, outline=(255, 255, 255, 60), width=8)
    return Image.alpha_composite(out, edge)


if __name__ == "__main__":
    big = build()
    big.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, "icon_preview.png"))
    sizes = [16, 24, 32, 48, 64, 128, 256]
    big.save(os.path.join(HERE, "KompasRelink.ico"), sizes=[(s, s) for s in sizes])
    print("ok")
