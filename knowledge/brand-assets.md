# Brand Assets in PowerPoint

Het hele HAN design system inzetten in een deck: logo, iconen, foto's en de
signatuur-stripe. Dit gaat verder dan kleur en font (die zitten al in de template).

Alle assets staan in:
`assets/brand/`

> **Pad-tip:** dit pad is relatief aan de skill-map (waar deze skill geinstalleerd is,
> bijvoorbeeld `~/.claude/skills/presentation-engine`). Bouw het absolute pad op met
> de variabele hieronder. Lees assets via bash/python; Glob en sommige PowerShell-calls
> struikelen soms over spaties in paden.

```python
import os
SKILL  = os.path.expanduser("~/.claude/skills/presentation-engine")  # pas aan
ASSETS = os.path.join(SKILL, "assets/brand")
```

---

## Wat is er

| Map | Inhoud | Formaat |
|-----|--------|---------|
| `logos/` | han-logo-fuccia, -white, -black(.svg/.png), -black-small, favicon | png + 1 svg |
| `icons/` | 47 iconen in HAN-rood | **webp** |
| `icons-black/` | 12 iconen in zwart | **webp** |
| `photos/` | campus, classroom, han-arnhem, mech-engineering-1/2, student-portrait | jpg + webp |
| `backgrounds/` | teams-bg, teams-bg-red, teams-bg-white | png |

**Iconenlijst (rood, `icons/`):** aanmelden, arrow, bachelor, beroepsmogelijkheden,
chat, check, chevron, close, cost, decision, download, duits, duur-opleiding, earn,
email, engels, external-link, facebook, hamburger, heart, helpdesk, i-know, info,
instagram, kennismaken, klassegrootte, lesdagen, livestream-red-white,
livestream-white-red, location, matching, maybe, min, more, no-idea, opleidingstype,
pause, phone, play, projects, search, star, start-date, stop, student, study-load, twitter.

`icons-black/` heeft een subset: aanmelden, arrow, check, chevron, close, email,
external-link, hamburger, info, play, star, student.

---

## De webp-gotcha (lees dit eerst)

python-pptx kan **geen webp** plaatsen. De iconen en de meeste foto's zijn webp.
Converteer ze on-the-fly naar PNG met Pillow (Pillow met webp-support is aanwezig).
Cache de PNG's in een tijdelijke map zodat je niet dubbel converteert.

```python
import os
from PIL import Image

SKILL  = os.path.expanduser("~/.claude/skills/presentation-engine")  # pas aan
ASSETS = os.path.join(SKILL, "assets/brand")
_CACHE = ".asset-cache"  # png-conversies, naast je buildscript

def asset_png(relpath):
    """Geef een bruikbaar PNG-pad terug voor een asset.

    relpath bv. 'icons/check.webp' of 'photos/classroom.webp'.
    png/jpg worden ongewijzigd doorgegeven; webp wordt naar png geconverteerd.
    """
    src = os.path.join(ASSETS, relpath)
    if not relpath.lower().endswith(".webp"):
        return src
    os.makedirs(_CACHE, exist_ok=True)
    dst = os.path.join(_CACHE, relpath.replace("/", "_")[:-5] + ".png")
    if not os.path.exists(dst):
        Image.open(src).convert("RGBA").save(dst, "PNG")
    return dst
```

Gebruik: `asset_png("icons/check.webp")` geeft een PNG-pad dat python-pptx accepteert.

---

## Logo plaatsen

Kies de logovariant op achtergrondkleur:
- Witte/lichte slide: `han-logo-black.png` of `han-logo-fuccia.png`
- Donkere/zwarte slide: `han-logo-white.png`

De template-layouts hebben zelf al HAN-branding in de master. Voeg een logo dus
alleen toe op slides die je volledig custom opbouwt (bijv. een full-bleed
afbeeldingsslide), niet op gewone placeholder-slides (dubbel logo = rommelig).

```python
from pptx.util import Cm

def add_logo(slide, variant="black", height=Cm(1.2), margin=Cm(0.8)):
    """Plaats het HAN-logo rechtsboven. variant: 'black' | 'white' | 'fuccia'."""
    path = asset_png(f"logos/han-logo-{variant}.png")
    pic = slide.shapes.add_picture(path, 0, margin, height=height)
    pic.left = SLIDE_W - pic.width - margin   # rechts uitlijnen
    return pic
```

---

## Iconen plaatsen

Iconen ondersteunen, ze vertellen het verhaal niet. Gebruik ze consistent: kies
rood óf zwart per deck, niet door elkaar. Eén icoon per card/punt, niet stapelen.

```python
from pptx.util import Cm, Emu

def add_icon(slide, name, left, top, size=Cm(1.4), color="red"):
    """Plaats een HAN-icoon. color: 'red' (icons/) of 'black' (icons-black/)."""
    folder = "icons" if color == "red" else "icons-black"
    path = asset_png(f"{folder}/{name}.webp")
    return slide.shapes.add_picture(path, left, top, height=size)
```

Veelgebruikt in slides: `check` (afvinken), `arrow`/`chevron` (volgende stap),
`star` (highlight), `chat`/`email`/`phone` (contact), `location`, `info`,
`decision`, `projects`, `student`.

**Icoon boven een card-titel** is een sterk patroon, sterker dan een bullet:

```python
# icoon gecentreerd boven een kolom, titel eronder
icon_size = Cm(1.6)
add_icon(slide, "decision", Emu(col_left + (col_w - icon_size)//2), col_top,
         size=icon_size, color="red")
add_textbox(slide, col_left, Emu(col_top + icon_size + Cm(0.3)), col_w, Cm(1),
            "Kiezen", font_size=Pt(18), bold=True, alignment=PP_ALIGN.CENTER)
```

---

## De signatuur-stripe (-12 graden)

Het meest herkenbare HAN-element: een schuine balk onder -12 graden, verhouding
breedte:hoogte ongeveer 6:1, in HAN-rood. **Eén per compositie.** Gebruik 'm op
de titeldia of een statement-slide als accent, nooit als versiering op elke slide.

Er is geen stripe-bestand; je tekent 'm als geroteerde rechthoek.

> **Let op botsing:** de stripe is een vlak, dus controleer dat hij geen
> placeholder-tekst (titel, ondertitel) of het master-logo overlapt. Op de titeldia
> is de lege rechterbovenhoek een veilige plek; op een statement-slide de onderrand.

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm

def add_stripe(slide, left, top, width=Cm(12), color=None):
    """HAN signatuur-stripe: schuine rode balk onder -12 graden, ratio 6:1."""
    height = Emu(width // 6)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    bar.rotation = -12
    bar.fill.solid()
    bar.fill.fore_color.rgb = color or HAN_RED
    bar.line.fill.background()
    bar.shadow.inherit = False
    return bar
```

---

## Foto's: behandeling

- Gebruik layout 3 (foto in placeholder) voor tekst + beeld naast elkaar.
- Voor full-bleed: plaats de foto eerst, dan tekst eroverheen in een halfdoorzichtig
  vlak. Houd contrast hoog (witte tekst op donkere foto, of een donkere overlay).
- Altijd via `asset_png()` (webp converteren).
- Geen pixelig/uitgerekt beeld: respecteer de verhouding, vul met `insert_picture`
  in de placeholder die de crop zelf regelt.

---

## Wat je NIET doet (uit han-design-system skill)

- Geen Material/Lucide/Heroicons. Alleen de HAN-iconenset.
- Geen gradients, glas, schaduwen-als-versiering, afgeronde hoeken boven 4px.
- Geen tweede accentkleur naast rood op tekst. Kleur zit in shapes, niet in losse tekst.
- Stripe en logo: spaarzaam. Eén stripe per compositie, logo niet dubbel met de master.

Volledige merkspec: de los te installeren `han-design-system` skill.
