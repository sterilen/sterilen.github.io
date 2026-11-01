import sys
from PIL import Image, ImageDraw, ImageFont

ut = sys.argv[1]
W, H = 460, 300
MORK, GRON, RAM, DAMP, LJUS = '#18222f', '#7fb046', '#0f1720', '#6b7280', '#9fb0c3'

def font(namn, storlek):
    for kand in (namn, 'segoeui.ttf', 'arial.ttf', 'DejaVuSans.ttf'):
        try:
            return ImageFont.truetype(kand, storlek)
        except Exception:
            pass
    return ImageFont.load_default()

SS = 2  # rita i dubbel storlek och skala ner, ger mjuka kanter
im = Image.new('RGB', (W * SS, H * SS), RAM)
d = ImageDraw.Draw(im)
d.rectangle([SS, SS, (W - 1) * SS, (H - 1) * SS], fill=MORK)
d.rectangle([SS, SS, (W - 1) * SS, 5 * SS], fill=GRON)
cx, cy, r = W // 2, 88, 35
d.ellipse([(cx - r) * SS, (cy - r) * SS, (cx + r) * SS, (cy + r) * SS], fill=GRON)

def mitt(text, y, f, farg):
    b = d.textbbox((0, 0), text, font=f)
    d.text(((W * SS - (b[2] - b[0])) // 2 - b[0], y * SS - b[1]), text, font=f, fill=farg)

mitt('L', cy - 20, font('segoeuib.ttf', 28 * SS), 'white')
mitt('Lagersök Admin', 140, font('segoeuib.ttf', 26 * SS), 'white')
mitt('Databasuppdatering v1.0', 182, font('segoeui.ttf', 11 * SS), LJUS)
d.rectangle([30 * SS, 232 * SS, (W - 30) * SS, 236 * SS], fill='#1e2a38')
mitt('Startar...', 252, font('segoeui.ttf', 9 * SS), DAMP)
im.resize((W, H), Image.LANCZOS).save(ut)
