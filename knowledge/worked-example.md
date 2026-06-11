# Worked Example: een compleet, gevarieerd deck

Dit is een end-to-end voorbeeld dat laat zien hoe een goed HAN-deck eruitziet:
afwisselende patronen, geen bullet-woestijn. Gebruik het als blauwdruk en
vervang de inhoud.

Het deck (8 slides) over "AI in het DM-team":
1. Titeldia (met stripe)
2. Statement (prikkelende vraag)
3. Afbeelding-slide (foto + tekst)
4. Card-rij (3 pijlers, met iconen)
5. Stappenflow (3 stappen)
6. Big numbers (statistieken)
7. Quote-slide (layout 5)
8. Statement + CTA (afsluiting)

Hier zónder kale bullet-slide, om de patronen te tonen. In een echt deck mag er
gerust een bullet-slide tussen zitten waar dat de beste vorm is, alleen niet overal.

---

## Het script

Sla op als `build_deck.py` in de projectmap van de presentatie en draai met `python build_deck.py`.
De volledige helper-bibliotheek staat in [design-components.md](design-components.md);
hieronder staan de helpers die dit voorbeeld nodig heeft inline, zodat het los draait.

```python
import os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ---- paden ----
# SKILL = absoluut pad naar deze skill-map. Pas aan naar waar je de skill installeerde,
# bijvoorbeeld ~/.claude/skills/presentation-engine
SKILL    = os.path.expanduser("~/.claude/skills/presentation-engine")
TEMPLATE = os.path.join(SKILL, "assets/template/han-powerpoint-template.pptx")
ASSETS   = os.path.join(SKILL, "assets/brand")
OUT      = "ai-in-het-dm-team.pptx"
_CACHE   = ".asset-cache"

# ---- kleuren ----
HAN_RED, HAN_BLACK, HAN_WHITE = RGBColor(0xE5,0x00,0x56), RGBColor(0,0,0), RGBColor(0xFF,0xFF,0xFF)
HAN_ALABASTER, HAN_MERCURY, HAN_GRAY = RGBColor(0xF8,0xF8,0xF8), RGBColor(0xE3,0xE3,0xE3), RGBColor(0x91,0x91,0x91)
HAN_MINESHAFT = RGBColor(0x45,0x45,0x45)

# ---- slide-grid ----
SLIDE_W, SLIDE_H = Cm(33.87), Cm(19.05)
MARGIN_LEFT = Cm(2.5)
CONTENT_W = Cm(28.87)
CONTENT_TOP, CONTENT_H = Cm(5.3), Cm(11.8)
GAP = Cm(0.8)

# ---- asset helper (webp -> png) ----
def asset_png(relpath):
    src = os.path.join(ASSETS, relpath)
    if not relpath.lower().endswith(".webp"):
        return src
    os.makedirs(_CACHE, exist_ok=True)
    dst = os.path.join(_CACHE, relpath.replace("/", "_")[:-5] + ".png")
    if not os.path.exists(dst):
        Image.open(src).convert("RGBA").save(dst, "PNG")
    return dst

# ---- mini-helpers ----
def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def add_textbox(slide, left, top, width, height, text, size=Pt(14),
                color=HAN_BLACK, bold=False, align=PP_ALIGN.LEFT):
    tf = slide.shapes.add_textbox(left, top, width, height).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text, p.font.size, p.font.bold = text, size, bold
    p.font.color.rgb, p.alignment = color, align
    return tf

def add_stripe(slide, left, top, width=Cm(12), color=HAN_RED):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Emu(width//6))
    bar.rotation = -12
    bar.fill.solid(); bar.fill.fore_color.rgb = color
    bar.line.fill.background(); bar.shadow.inherit = False
    return bar

def add_icon(slide, name, left, top, size=Cm(1.4), color="red"):
    folder = "icons" if color == "red" else "icons-black"
    return slide.shapes.add_picture(asset_png(f"{folder}/{name}.webp"), left, top, height=size)

def distribute(n, item_w, gap=GAP):
    total = item_w*n + gap*(n-1)
    start = Emu(MARGIN_LEFT + (CONTENT_W - total)//2)
    return [Emu(start + (item_w+gap)*i) for i in range(n)]

def add_card(slide, left, top, w, h, title, body, border=HAN_RED, align=PP_ALIGN.CENTER):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    s.adjustments[0] = 0.04
    s.fill.solid(); s.fill.fore_color.rgb = HAN_WHITE
    s.line.color.rgb = border; s.line.width = Pt(2); s.shadow.inherit = False
    tf = s.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = title; p.font.bold = True
    p.font.size = Pt(18); p.font.color.rgb = HAN_BLACK; p.alignment = align
    if body:
        p2 = tf.add_paragraph(); p2.text = body; p2.font.size = Pt(12)
        p2.font.color.rgb = HAN_MINESHAFT; p2.alignment = align; p2.space_before = Pt(6)
    return s

# ---- template laden + voorbeeldslides weg ----
prs = Presentation(TEMPLATE)
while len(prs.slides._sldIdLst) > 0:
    rId = prs.slides._sldIdLst[0].get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    prs.part.drop_rel(rId); prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

# 1. TITELDIA + stripe
s = prs.slides.add_slide(prs.slide_layouts[0])
s.placeholders[0].text  = "AI in het DM-team"
s.placeholders[14].text = "Werkoverleg Digital Marketing"
s.placeholders[13].text = "Guus Witjes"
# Stripe als accent in de lege rechterbovenhoek, weg van titel/ondertitel/logo
add_stripe(s, Cm(22.5), Cm(2.2), width=Cm(10))
add_notes(s, "Korte opener. Niet voorlezen, gewoon de toon zetten.")

# 2. STATEMENT (layout 2, half width)
s = prs.slides.add_slide(prs.slide_layouts[2])
s.placeholders[0].text = "Wat als elk teamlid zijn eigen AI-assistent had?"
add_notes(s, "Stel de vraag. Wacht. Laat het even hangen voor je doorgaat.")

# 3. AFBEELDING (layout 3)
s = prs.slides.add_slide(prs.slide_layouts[3])
s.placeholders[0].text  = "Geen losse trucs, maar een werkwijze"
s.placeholders[12].text = "AI wordt pas waardevol als het in je dagelijkse werk zit, niet ernaast."
s.placeholders[11].insert_picture(asset_png("photos/classroom.webp"))
add_notes(s, "Vertel het concrete verhaal achter de foto.")

# 4. CARD-RIJ met iconen (3 pijlers)
s = prs.slides.add_slide(prs.slide_layouts[1])
s.placeholders[0].text = "Drie pijlers"
if s.placeholders[11].has_text_frame: s.placeholders[11].text = ""
pijlers = [("decision","Kiezen","Waar voegt AI echt waarde toe?"),
           ("projects","Bouwen","Werkwijzen, niet losse prompts."),
           ("star","Delen","Wat werkt, gaat het team in.")]
card_w = Emu((CONTENT_W - GAP*2)//3); card_h = Cm(6)
xs = distribute(3, card_w); top = Cm(8)
for x,(icon,titel,body) in zip(xs, pijlers):
    isz = Cm(1.6)
    add_icon(s, icon, Emu(x + (card_w-isz)//2), Emu(top - isz - Cm(0.3)), size=isz)
    add_card(s, x, top, card_w, card_h, titel, body)
add_notes(s, "Per pijler een voorbeeld uit het team.")

# 5. STAPPENFLOW (3 stappen met pijlen)
s = prs.slides.add_slide(prs.slide_layouts[1])
s.placeholders[0].text = "Hoe we het aanpakken"
if s.placeholders[11].has_text_frame: s.placeholders[11].text = ""
stappen = [("1","Verkennen"),("2","Bouwen"),("3","Borgen")]
step_w, arrow_w = Cm(7), Cm(1.2); top = Cm(9)
total = step_w*3 + arrow_w*2; start = Emu(MARGIN_LEFT + (CONTENT_W-total)//2)
for i,(num,label) in enumerate(stappen):
    left = Emu(start + (step_w+arrow_w)*i)
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, Emu(left+(step_w-Cm(1.8))//2), top, Cm(1.8), Cm(1.8))
    badge.fill.solid(); badge.fill.fore_color.rgb = HAN_RED; badge.line.fill.background()
    badge.shadow.inherit = False
    bt = badge.text_frame.paragraphs[0]; bt.text = num; bt.font.bold = True
    bt.font.size = Pt(20); bt.font.color.rgb = HAN_WHITE; bt.alignment = PP_ALIGN.CENTER
    badge.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_textbox(s, left, Emu(top+Cm(2.0)), step_w, Cm(1), label, size=Pt(18), bold=True, align=PP_ALIGN.CENTER)
    if i < 2:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(left+step_w), Emu(top+Cm(0.5)), arrow_w, Cm(0.8))
        ar.fill.solid(); ar.fill.fore_color.rgb = HAN_GRAY; ar.line.fill.background(); ar.shadow.inherit=False
add_notes(s, "Benadruk dat borgen het verschil maakt, niet de pilot.")

# 6. BIG NUMBERS
s = prs.slides.add_slide(prs.slide_layouts[1])
s.placeholders[0].text = "Waar staan we"
if s.placeholders[11].has_text_frame: s.placeholders[11].text = ""
stats = [("8","teamleden aan de slag"),("3","werkwijzen live"),("40%","minder routinewerk")]
item_w, card_h = Cm(8), Cm(6.5); xs = distribute(3, item_w); top = Cm(8)
for x,(num,label) in zip(xs, stats):
    c = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top, item_w, card_h)
    c.adjustments[0]=0.04; c.fill.solid(); c.fill.fore_color.rgb=HAN_WHITE
    c.line.color.rgb=HAN_RED; c.line.width=Pt(2); c.shadow.inherit=False
    tf=c.text_frame; tf.vertical_anchor=MSO_ANCHOR.MIDDLE; tf.word_wrap=True
    p=tf.paragraphs[0]; p.text=num; p.font.size=Pt(54); p.font.bold=True
    p.font.color.rgb=HAN_BLACK; p.alignment=PP_ALIGN.CENTER
    p2=tf.add_paragraph(); p2.text=label; p2.font.size=Pt(14)
    p2.font.color.rgb=HAN_MINESHAFT; p2.alignment=PP_ALIGN.CENTER; p2.space_before=Pt(4)
add_notes(s, "Cijfers zijn indicatief, vervang door echte teamcijfers.")

# 7. QUOTE (layout 5)
s = prs.slides.add_slide(prs.slide_layouts[5])
s.placeholders[0].text  = "AI is geen tool die je erbij pakt. Het is een manier van werken."
s.placeholders[11].text = "Guus Witjes"
add_notes(s, "Laat de quote staan. Stilte werkt hier.")

# 8. STATEMENT + CTA
s = prs.slides.add_slide(prs.slide_layouts[2])
s.placeholders[0].text = "Volgende stap: kies één werkwijze die je deze maand bouwt."
add_notes(s, "Concreet maken. Wie pakt wat op? Sluit af met een afspraak.")

prs.save(OUT)
print("Gebouwd:", OUT, "|", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
```

---

## Wat dit voorbeeld goed doet

- **Opent met de Titeldia + stripe.** Niet met een agenda-lijst.
- **Acht slides, acht verschillende patronen.** Geen herhaling, geen kale bullets.
- **Iconen boven cards** in plaats van bullet-punten.
- **Kleur zit in shapes** (card-borders, badges), niet in losse tekst.
- **Elke slide heeft speaker notes** met de details die niet op de slide staan.
- **Sluit af met een actie**, niet met "Bedankt voor de aandacht".

## Daarna: fonts insluiten

Wil je de deck delen met iemand zonder Roboto? Draai na het bouwen:

```
powershell -ExecutionPolicy Bypass -File scripts/embed_fonts.ps1 -Path "ai-in-het-dm-team.pptx"
```
