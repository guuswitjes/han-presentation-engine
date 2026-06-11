# HAN PowerPoint Template — Layout Referentie

Template locatie: `assets/template/han-powerpoint-template.pptx`

---

## Titel-stijl

Alle layouts (behalve Titeldia en Vergelijking) hebben dezelfde titel-placeholder:
- **Positie:** left=2.3cm, top=1.0cm, width=29.2cm, height=3.7cm
- **Stijl:** HAN Rood, hoofdletters (via theme)
- **Gebruik altijd `slide.placeholders[0].text = "..."`** voor de titel. Nooit een custom textbox.

---

## Layouts

### Layout 0: Titeldia
Gebruik: openingsslide van de presentatie.

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Hoofdtitel |
| 14 | BODY | Boventekst (bijv. datum, context) |
| 13 | BODY | Ondertitel |

### Layout 1: Titel en Tekst (full width)
Gebruik: standaard contentslide met tekst over de volle breedte.

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Slide titel |
| 11 | BODY | Tekst (full width) |
| 12 | SLIDE_NUMBER | Automatisch |

### Layout 2: Titel en Tekst (half width)
Gebruik: contentslide met tekst links (halve breedte), rechts open voor vrije ruimte.

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Slide titel |
| 11 | BODY | Tekst (linkerhelft) |
| 12 | SLIDE_NUMBER | Automatisch |

### Layout 3: Titel, Tekst en Afbeelding
Gebruik: contentslide met tekst links en afbeelding rechts.

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Slide titel |
| 12 | BODY | Tekst (linkerhelft) |
| 11 | PICTURE | Afbeelding (rechterhelft) |
| 13 | SLIDE_NUMBER | Automatisch |

### Layout 4: Dubbele Titel en Tekst
Gebruik: twee kolommen met elk een subtitel en tekst. Goed voor vergelijkingen of twee onderwerpen.

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Hoofdtitel |
| 16 | BODY | Subtitel links |
| 15 | BODY | Subtitel rechts |
| 18 | BODY | Tekst links |
| 19 | BODY | Tekst rechts |
| 20 | SLIDE_NUMBER | Automatisch |

### Layout 5: Vergelijking / Quote
Gebruik: quote-slide of vergelijking. Heeft een groot zwart vlak (rechterhelft) met daarin de quote-tekst en een naambalk eronder. Linkerhelft blijft open (bijv. voor een foto of leeg).

**Visuele structuur:**
- Zwart vlak: 16.4 x 15 cm, start op 8.7cm links, 2cm top
- Klein aanhalingsteken-icoon linksboven in het zwarte vlak
- Quote-tekst (wit op zwart) in het grote TITLE-veld
- Naambalk onderaan het zwarte vlak

| Placeholder idx | Type | Gebruik |
|-----------------|------|---------|
| 0 | TITLE | Quote tekst (wit op zwart vlak) |
| 11 | BODY | Naam / bron onder de quote |
| 13 | SLIDE_NUMBER | Automatisch |

**Let op:** het zwarte vlak en aanhalingsteken-icoon zitten in de layout zelf, die hoef je niet te bouwen. Vul alleen de placeholders.

---

## Python-pptx Bouwcode

### Template laden en voorbeeldslides verwijderen

```python
from pptx import Presentation
from pptx.util import Inches, Pt

TEMPLATE_PATH = 'assets/template/han-powerpoint-template.pptx'

prs = Presentation(TEMPLATE_PATH)

# Verwijder alle voorbeeldslides uit de template
while len(prs.slides._sldIdLst) > 0:
    rId = prs.slides._sldIdLst[0].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    )
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
```

### Slide toevoegen

```python
# Titelslide
layout = prs.slide_layouts[0]  # Titeldia
slide = prs.slides.add_slide(layout)
slide.placeholders[0].text = "Presentatie Titel"
slide.placeholders[13].text = "Ondertitel"
slide.placeholders[14].text = "Datum of context"

# Contentslide (full width)
layout = prs.slide_layouts[1]  # Titel en Tekst_
slide = prs.slides.add_slide(layout)
slide.placeholders[0].text = "Slide Titel"
slide.placeholders[11].text = "Bullet punt 1\nBullet punt 2"

# Slide met afbeelding
layout = prs.slide_layouts[3]  # Titel, Tekst en Afbeelding
slide = prs.slides.add_slide(layout)
slide.placeholders[0].text = "Slide Titel"
slide.placeholders[12].text = "Tekst naast de afbeelding"
slide.placeholders[11].insert_picture('pad/naar/afbeelding.png')
```

### Speaker notes toevoegen

```python
from pptx.util import Pt

notes_slide = slide.notes_slide
notes_slide.notes_text_frame.text = "Speaker notes tekst hier."
```

### Opslaan

```python
prs.save('output/presentatie-naam.pptx')
```

---

## Veelgemaakte fouten

1. **Verkeerde placeholder idx** — Altijd checken via de tabel hierboven. De idx verschilt per layout
2. **Tekst overschrijft formatting** — Gebruik `.text =` voor eenvoudige tekst. Voor opmaak, werk met `paragraphs` en `runs`
3. **Afbeelding in verkeerde placeholder** — Layout 3 heeft de picture placeholder op idx 11, niet 12
4. **Voorbeeldslides niet verwijderd** — Template bevat 6 voorbeeldslides die je moet verwijderen voor je begint
