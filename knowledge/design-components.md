# Design Components Library

Herbruikbare python-pptx bouwblokken voor visueel rijke slides. Gebruik deze in plaats van (of naast) template placeholders wanneer een slide meer nodig heeft dan titel + tekst.

---

## HAN Kleuren (python-pptx RGBColor)

```python
from pptx.dml.color import RGBColor

# Primair
HAN_RED    = RGBColor(0xE5, 0x00, 0x56)
HAN_BLACK  = RGBColor(0x00, 0x00, 0x00)
HAN_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

# Secundair
HAN_ORANGE = RGBColor(0xF5, 0xA8, 0x46)
HAN_GREEN  = RGBColor(0x7A, 0xC4, 0x8E)
HAN_BLUE   = RGBColor(0x6A, 0xC3, 0xED)

# Tertiair
HAN_OLIVE  = RGBColor(0xBB, 0xD7, 0x37)
HAN_YELLOW = RGBColor(0xF8, 0xE2, 0x60)
HAN_CORAL  = RGBColor(0xEB, 0x52, 0x41)
HAN_VIOLET = RGBColor(0xB1, 0x41, 0x96)
HAN_NAVY   = RGBColor(0x32, 0x5C, 0xA7)

# Grijs
HAN_ALABASTER = RGBColor(0xF8, 0xF8, 0xF8)
HAN_MERCURY   = RGBColor(0xE3, 0xE3, 0xE3)
HAN_GRAY      = RGBColor(0x91, 0x91, 0x91)
HAN_BOULDER   = RGBColor(0x75, 0x75, 0x75)
HAN_MINESHAFT = RGBColor(0x45, 0x45, 0x45)
```

---

## Basis imports

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
```

---

## Slide constanten en layout grid

De HAN template is 33.87 x 19.05 cm (widescreen 16:9). Alle componenten gebruiken deze constanten voor consistente uitlijning.

```python
# Slide dimensies
SLIDE_W = Cm(33.87)
SLIDE_H = Cm(19.05)

# Marges
MARGIN_LEFT   = Cm(2.5)
MARGIN_RIGHT  = Cm(2.5)
MARGIN_TOP    = Cm(1.5)
MARGIN_BOTTOM = Cm(1.5)

# Beschikbare content-ruimte
CONTENT_W = Cm(28.87)   # SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_TOP = Cm(4.5)   # Na slide titel
CONTENT_BOTTOM = Cm(17.5)  # SLIDE_H - MARGIN_BOTTOM
CONTENT_H = Cm(13)      # CONTENT_BOTTOM - CONTENT_TOP

# Standaard gap tussen elementen
GAP = Cm(0.8)
GAP_SMALL = Cm(0.4)
```

---

## Uitlijningsregels

**KRITIEK: deze regels gelden voor ALLE componenten. Niet handmatig positioneren.**

### Verticale uitlijning van elementen naast elkaar

Wanneer een cirkel-badge naast tekst staat, of een card naast een pijl:
- Bereken het **verticale midden** van het hoogste element
- Lijn alle andere elementen uit op dat midden
- Formule: `element_top = reference_center - (element_height / 2)`

```python
def align_middle(reference_top, reference_height, element_height):
    """Bereken top-positie zodat element verticaal gecentreerd is t.o.v. reference."""
    reference_center = reference_top + reference_height // 2
    return reference_center - element_height // 2
```

### Horizontaal verdelen van elementen

Wanneer meerdere cards of elementen naast elkaar staan:
- Bereken de totale breedte van alle elementen + gaps
- Centreer het geheel op de slide
- Formule: `start_left = MARGIN_LEFT + (CONTENT_W - total_width) / 2`

```python
def distribute_horizontal(n_items, item_width, gap=GAP):
    """Bereken startposities voor n gelijk verdeelde items, gecentreerd op de slide."""
    total = item_width * n_items + gap * (n_items - 1)
    start = Emu(MARGIN_LEFT + (CONTENT_W - total) // 2)
    return [Emu(start + (item_width + gap) * i) for i in range(n_items)]
```

### Verticaal centreren van content-blok

Wanneer je totale content (cards + bars + tekst) kleiner is dan de beschikbare ruimte, centreer het blok verticaal zodat de slide gevuld aanvoelt.

```python
def center_content_top(content_height):
    """Bereken top-positie zodat content verticaal gecentreerd is in het content-gebied.

    Gebruik dit als startpositie in plaats van CONTENT_TOP wanneer de content
    minder dan ~70% van CONTENT_H vult.
    """
    return Emu(CONTENT_TOP + (CONTENT_H - content_height) // 2)
```

**Wanneer gebruiken:**
- Content hoogte < 70% van CONTENT_H (~9 cm): centreer verticaal
- Content hoogte >= 70%: gebruik CONTENT_TOP als startpunt (genoeg vulling)

### Rij-uitlijning (cirkel + tekst + card)

Het meest voorkomende patroon: cirkel-badge links, tekst in het midden, card rechts. Allemaal op dezelfde basislijn.

```python
def calculate_row_positions(row_top, row_height, badge_size, text_height, card_height):
    """Bereken verticale posities zodat badge, tekst en card gecentreerd zijn op row_height."""
    center = row_top + row_height // 2
    return {
        'badge_top': center - badge_size // 2,
        'text_top': center - text_height // 2,
        'card_top': center - card_height // 2,
    }
```

---

## Component 1: Card (afgerond rechthoek met tekst)

```python
def add_card(slide, left, top, width, height, title, body,
             border_color=HAN_RED, fill_color=HAN_WHITE,
             title_size=Pt(16), body_size=Pt(12), title_color=HAN_BLACK,
             alignment=PP_ALIGN.LEFT):
    """Voegt een card toe met afgeronde hoeken, border en tekst.

    alignment geldt voor ALLE tekst in de card (titel en body).
    Standaard links. Gebruik PP_ALIGN.LEFT of PP_ALIGN.CENTER, niet mixen.
    """
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.adjustments[0] = 0.05
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(2)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.8)
    tf.margin_right = Cm(0.8)
    tf.margin_top = Cm(0.6)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # tekst verticaal gecentreerd in card

    p = tf.paragraphs[0]
    p.text = title
    p.font.size = title_size
    p.font.bold = True
    p.font.color.rgb = title_color
    p.alignment = alignment

    if body:
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.size = body_size
        p2.font.color.rgb = HAN_MINESHAFT
        p2.space_before = Pt(8)
        p2.alignment = alignment

    return shape
```

---

## Component 2: Card Row (automatisch gecentreerde rij van cards)

Vervangt handmatig positioneren van meerdere cards naast elkaar.

```python
def add_card_row(slide, cards, top=None, card_height=Cm(5), gap=GAP):
    """Rij van cards, automatisch gecentreerd op de slide.

    cards = [(title, body, border_color, title_color), ...]
    top = None: automatisch verticaal gecentreerd. Geef waarde mee om te overriden.
    """
    n = len(cards)
    card_w = Emu((CONTENT_W - gap * (n - 1)) // n)
    positions = distribute_horizontal(n, card_w, gap)

    if top is None:
        top = center_content_top(card_height)

    shapes = []
    for i, (title, body, border_color, title_color) in enumerate(cards):
        shape = add_card(slide, positions[i], top, card_w, card_height,
                        title, body,
                        border_color=border_color, title_color=title_color,
                        title_size=Pt(20), body_size=Pt(13))
        shapes.append(shape)

    return shapes
```

---

## Component 3: Vergelijkingsblok (twee cards met pijlen)

Gecentreerd, met chevron-pijlen ertussen.

```python
def add_comparison(slide, left_title, left_body, left_color,
                   right_title, right_body, right_color,
                   top=None):
    """Twee cards naast elkaar met chevron-pijlen ertussen, gecentreerd.

    Cards worden verticaal gecentreerd. Tekst in cards is verticaal
    gecentreerd via MSO_ANCHOR.MIDDLE. Pijlen hebben voldoende ademruimte.
    """
    card_w = Cm(11.5)
    card_h = Cm(5)       # compact: passend bij 2-4 regels tekst
    arrow_zone = Cm(5)   # ruimte voor 3 chevrons + ademruimte
    total = card_w * 2 + arrow_zone
    start = Emu(MARGIN_LEFT + (CONTENT_W - total) // 2)
    if top is None:
        top = center_content_top(card_h)

    # Linker card
    add_card(slide, start, top, card_w, card_h,
             left_title, left_body,
             border_color=left_color, title_color=left_color,
             title_size=Pt(20), body_size=Pt(14))

    # Chevron-pijlen (verticaal gecentreerd op cards, met ruimte aan weerszijden)
    arrow_h = Cm(1.2)
    arrow_top = align_middle(top, card_h, arrow_h)
    arrow_start = Emu(start + card_w + Cm(1))  # 1 cm ademruimte na card
    for i in range(3):
        arrow = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON,
            Emu(arrow_start + Cm(1.2) * i), arrow_top,
            Cm(1), arrow_h
        )
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = HAN_GRAY
        arrow.line.fill.background()

    # Rechter card (1 cm ademruimte na pijlen)
    add_card(slide, Emu(start + card_w + arrow_zone), top, card_w, card_h,
             right_title, right_body,
             border_color=right_color, title_color=right_color,
             title_size=Pt(20), body_size=Pt(14))
```

---

## Component 4: Labeled Row (cirkel-badge + tekst + card op een rij)

Het meest gebruikte patroon. Garandeert verticale uitlijning.

```python
def add_labeled_row(slide, row_top, number, label, detail,
                    color=HAN_RED, badge_size=Cm(2),
                    label_width=Cm(8), card_width=None):
    """Cirkel-badge + label + detail-card, verticaal uitgelijnd op middellijn.

    row_top: bovenkant van de rij
    number: tekst in de cirkel (bijv. "1")
    label: tekst naast de cirkel
    detail: tekst in de card rechts
    """
    row_h = Cm(2.5)  # standaard rijhoogte
    if card_width is None:
        card_width = Emu(CONTENT_W - MARGIN_LEFT - badge_size - GAP - label_width - GAP)

    # Posities berekenen (alles gecentreerd op row_h)
    pos = calculate_row_positions(row_top, row_h, badge_size, row_h, row_h)

    # Badge
    badge_left = MARGIN_LEFT
    add_circle_badge(slide, badge_left, pos['badge_top'], badge_size, number,
                     fill_color=color, border_color=color, text_color=HAN_WHITE)

    # Label
    label_left = Emu(badge_left + badge_size + GAP)
    add_textbox(slide, label_left, pos['text_top'], label_width, row_h,
                label, font_size=Pt(16), bold=True)

    # Detail card
    card_left = Emu(label_left + label_width + GAP)
    add_card(slide, card_left, pos['card_top'], card_width, row_h,
             detail, None,
             border_color=color, fill_color=HAN_ALABASTER,
             title_size=Pt(11), title_color=HAN_MINESHAFT)
```

---

## Component 5: Step Flow (genummerde stappen met pijlen)

Horizontale rij van genummerde stappen, gecentreerd.

```python
def add_step_flow(slide, steps, top=None, color=HAN_RED):
    """Rij van genummerde stappen met pijlen ertussen.

    steps = [(number, label, detail), ...]
    color: enkele kleur voor alle badges. Gebruik een enkele kleur tenzij er een
           inhoudelijke reden is voor kleurverschil (bijv. twee duidelijke fasen).
           Meerdere kleuren zonder logica = visuele ruis.
    top = None: automatisch verticaal gecentreerd.
    """
    n = len(steps)
    step_w = Cm(5.5)
    arrow_w = Cm(0.7)
    total = step_w * n + arrow_w * (n - 1)
    start = Emu(MARGIN_LEFT + (CONTENT_W - total) // 2)
    total_h = Cm(1.8) + GAP_SMALL + Cm(1.8) + Cm(2.5)  # badge + label + card
    if top is None:
        top = center_content_top(total_h)

    badge_size = Cm(1.8)

    for i, (number, label, detail) in enumerate(steps):
        left = Emu(start + (step_w + arrow_w) * i)

        # Badge (gecentreerd boven step_w)
        badge_left = Emu(left + (step_w - badge_size) // 2)
        add_circle_badge(slide, badge_left, top, badge_size, number,
                         fill_color=color, border_color=color,
                         text_color=HAN_WHITE)

        # Label (gecentreerd onder badge)
        label_top = Emu(top + badge_size + GAP_SMALL)
        add_textbox(slide, left, label_top, step_w, Cm(1.5),
                    label, font_size=Pt(16), bold=True,
                    alignment=PP_ALIGN.CENTER)

        # Detail card (gecentreerd onder label)
        card_top = Emu(label_top + Cm(1.8))
        add_card(slide, left, card_top, step_w, Cm(2.5),
                 detail, None,
                 border_color=HAN_MERCURY, fill_color=HAN_ALABASTER,
                 title_size=Pt(11), title_color=HAN_MINESHAFT,
                 alignment=PP_ALIGN.CENTER)

        # Pijl naar volgende stap (verticaal gecentreerd op badge)
        if i < n - 1:
            arrow_left = Emu(left + step_w)
            arrow_top = align_middle(top, badge_size, Cm(0.8))
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW, arrow_left, arrow_top, arrow_w, Cm(0.8)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = HAN_GRAY
            arrow.line.fill.background()
```

---

## Component 6: Tekstbox (vrij geplaatst)

> **Geen accent bars.** De HAN template heeft geen gekleurde balken. Gebruik cards met gekleurde borders als je kleuraccenten wilt.

```python
def add_textbox(slide, left, top, width, height, text,
                font_size=Pt(14), color=HAN_BLACK, bold=False,
                alignment=PP_ALIGN.LEFT):
    """Vrij geplaatste tekstbox zonder achtergrond."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.color.rgb = color
    p.font.bold = bold
    p.alignment = alignment
    return txBox
```

---

## Component 7: Cirkel-badge

```python
def add_circle_badge(slide, left, top, size, text,
                     fill_color=None, border_color=HAN_RED,
                     text_color=HAN_BLACK):
    """Cirkel met tekst erin (bijv. letter of getal)."""
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left, top, size, size
    )
    if fill_color:
        circle.fill.solid()
        circle.fill.fore_color.rgb = fill_color
    else:
        circle.fill.background()
    circle.line.color.rgb = border_color
    circle.line.width = Pt(2)

    tf = circle.text_frame
    tf.paragraphs[0].text = text
    tf.paragraphs[0].font.size = Pt(16)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = text_color
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    return circle
```

---

## Component 8: Slide titel

Gebruik altijd de **template placeholder** voor de titel (placeholder idx 0). Dit geeft automatisch de juiste kleur (HAN Rood), positie (left=2.3cm, top=1.0cm) en stijl.

```python
def add_slide_title(slide, text):
    """Stel de slide titel in via de template placeholder.

    Gebruikt placeholder idx 0 die in elke layout zit.
    Kleur, positie en font komen uit de template (HAN Rood, hoofdletters).
    NIET handmatig een textbox maken voor titels.
    """
    slide.placeholders[0].text = text
```

**Let op:** bij blank slides (layout 1) moet je NIET de placeholder tekst leegmaken voor idx 0 als je daarna `add_slide_title` wilt gebruiken. Maak alleen de body placeholder leeg (idx 11).

---

## Component 9: Big Number (statistiek)

```python
def add_big_numbers(slide, stats, top=None):
    """Rij van grote getallen in cards met labels, gecentreerd.

    stats = [(number, label, color), ...]
    Elk getal komt in een card met gekleurde border. Tekst zelf is zwart.
    Kleur zit in de shape, niet in de tekst (voorkomt circus-effect).
    top = None: automatisch verticaal gecentreerd.
    """
    n = len(stats)
    item_w = Cm(7)
    card_h = Cm(7)
    positions = distribute_horizontal(n, item_w, GAP)
    if top is None:
        top = center_content_top(card_h)

    for i, (number, label, color) in enumerate(stats):
        # Card met gekleurde border
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, positions[i], top, item_w, card_h
        )
        shape.adjustments[0] = 0.05
        shape.fill.solid()
        shape.fill.fore_color.rgb = HAN_WHITE
        shape.line.color.rgb = color
        shape.line.width = Pt(2)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Groot getal (zwart, niet gekleurd)
        p = tf.paragraphs[0]
        p.text = number
        p.font.size = Pt(56)
        p.font.bold = True
        p.font.color.rgb = HAN_BLACK
        p.alignment = PP_ALIGN.CENTER

        # Label
        p2 = tf.add_paragraph()
        p2.text = label
        p2.font.size = Pt(14)
        p2.font.color.rgb = HAN_MINESHAFT
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(4)
```

---

## Component 10: Speaker notes

```python
def add_notes(slide, text):
    """Voeg speaker notes toe aan een slide."""
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = text
```

---

## Wanneer welk component

| Slide type | Component | Functie |
|------------|-----------|---------|
| Vergelijking (voor/na) | `add_comparison` | 2 cards + chevrons, gecentreerd |
| Stappenplan / proces | `add_step_flow` | Genummerde stappen + pijlen |
| Genummerde lijst met details | `add_labeled_row` | Badge + label + card, uitgelijnd |
| Meerdere cards naast elkaar | `add_card_row` | Automatisch gecentreerd |
| Statistieken | `add_big_numbers` | Getallen in cards met gekleurde border |
| Quote / citaat | Template layout 5 (Vergelijking) | Zwart vlak met witte tekst, ingebouwd |
| Statement | `add_textbox` (groot, gecentreerd) | Vrij geplaatst |
| Kernpunt met toelichting | `add_card` | Vrij geplaatst |

---

## Design regels

### Kleurgebruik
- **Kleur hoort bij een visueel element** (card border, badge fill), niet bij losse tekst
- Meerdere kleuren op kale tekst ziet eruit als een circus. Dezelfde kleuren in shapes werken prima
- Losse tekst: gebruik alleen HAN_BLACK, HAN_MINESHAFT of HAN_BOULDER. Nooit accent-kleuren
- Als je meerdere kleuren wilt laten zien (bijv. statistieken): zet ze in cards met gekleurde borders
- Uitzondering: een enkele gekleurde titel-tekst is OK als accent, maar niet 4 naast elkaar
- **Meerdere kleuren in stappen/badges**: alleen als er een inhoudelijke reden is (bijv. twee fasen). Geen reden = een kleur voor alles

### Template-elementen boven custom shapes
- Gebruik **template layouts** als die het doel dekken (quote = layout 5, tekst+afbeelding = layout 3)
- Bouw alleen custom shapes als geen template layout past
- **Geen decoratieve elementen toevoegen** die niet in de template zitten (geen balken, strepen, lijnen als versiering)

### Card sizing
- **Card hoogte passend bij content**: 2 regels tekst = ~Cm(4), 3-4 regels = ~Cm(5), 5+ regels = Cm(6-7)
- **Nooit meer dan 50% lege ruimte in een card** - maak de card kleiner
- Tekst is verticaal gecentreerd in cards (via `MSO_ANCHOR.MIDDLE`)

### Tekstuitlijning in cards
- **Titel en body dezelfde uitlijning** - niet mixen (titel gecentreerd, body links)
- **Keuze gecentreerd vs links hangt af van de tekstlengte:**
  - **Gecentreerd**: korte tekst, max 1-3 woorden per regel (bijv. "Patroon herkenning", "SKILL.md + knowledge/")
  - **Links**: langere tekst, zinnen of beschrijvingen (bijv. "Triathlon, voeding, ontwikkeling\n\n1 agent, 4 skills")
- Niet de positie van omliggende elementen bepaalt dit, maar de tekst in de card zelf

### Ademruimte
- **Tussen cards en pijlen/andere elementen**: minimaal Cm(1)
- **Tussen cards onderling** (in een rij): `GAP` = Cm(0.8)
- **Marges**: altijd `MARGIN_LEFT`/`MARGIN_RIGHT` respecteren

### Verticale balans
- Content die minder dan 70% van de slide vult: verticaal centreren met `center_content_top()`
- Alle layout-componenten doen dit automatisch als `top=None`

## Tips

1. **Gebruik de layout-componenten** (`add_card_row`, `add_comparison`, `add_step_flow`, `add_labeled_row`) in plaats van handmatig positioneren
2. **Gebruik `align_middle()` en `distribute_horizontal()`** als je toch handmatig moet positioneren
3. **Gebruik max 2-3 kleuren per slide** met HAN Rood als primair accent
4. **Consistentie**: kies een component-stijl en houd die de hele presentatie vol
5. **Witruimte**: gebruik `MARGIN_LEFT`, `MARGIN_RIGHT`, `CONTENT_TOP` als startpunten
6. **Tekst in shapes**: max 2-3 regels per card, rest in speaker notes
