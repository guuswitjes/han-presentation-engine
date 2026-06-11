# Slide Patterns: van inhoud naar de juiste slide

Dit is de belangrijkste file van de skill. Hij lost het meest voorkomende probleem op:
**een deck dat alleen uit bullet-lijstjes bestaat.**

De template kan veel meer dan bullets: titeldia, quote-slide, afbeelding-slide,
twee-koloms, en custom shapes (cards, stappen, big numbers). Bullets zijn prima,
maar **niet als elke slide bullets is.** Deze file helpt je per slide bewust een
patroon te kiezen, zodat een deck afwisselt in plaats van te verzanden in lijstjes.

---

## De regel

> **Bullets mogen. Een deck dat alléén uit bullets bestaat, niet.**

Een bullet-lijst is een legitiem patroon: voor een echte opsomming van losse,
ongelijksoortige punten is het de juiste vorm. Het wordt pas een probleem als het
de standaard is voor elke slide. Voor je een body vult met `- punt 1 \n - punt 2`,
stel jezelf één vraag: *wat is de vórm van deze inhoud?* Vaak past een ander patroon
beter (stappen, cards, vergelijking). Gebruik de beslistabel hieronder.

### Harde grenzen (controleer dit op het hele deck)

- **Maximaal 1 op de 3 contentslides** is een kale bullet-body. Bullets zijn welkom, dominantie niet.
- **Nooit twee bullet-slides achter elkaar.** Wissel af.
- **Geen enkel deck zonder minstens één** van: quote/statement-slide, afbeelding-slide, of een visueel patroon (cards/stappen/vergelijking).
- **De openingsslide is altijd de Titeldia** (layout 0), nooit een bullet-slide met "Agenda" als eerste indruk.
- **Maximaal 6 regels per slide, streef naar 1-3.** De rest gaat naar speaker notes.

Als een deck deze grenzen overschrijdt: herontwerp de overtredende slides voor je opslaat.

---

## Beslistabel: welk patroon bij welke inhoud

| Wat wil de slide doen? | Patroon | Hoe (layout / component) |
|------------------------|---------|--------------------------|
| Openen, context zetten | **Titeldia** | layout 0 + eyebrow + stripe |
| Eén krachtig statement / cijfer laten landen | **Statement** | layout 2, grote tekst, veel wit |
| Een citaat tonen | **Quote-slide** | layout 5 (Vergelijking), zwart vlak |
| Emotie / herkenning / "zo ziet het eruit" | **Afbeelding** | layout 3, foto rechts |
| Voor naar na, oud naar nieuw, A vs B | **Vergelijking** | `add_comparison` of layout 4 |
| Een proces met stappen | **Stappenflow** | `add_step_flow` |
| 2-4 gelijkwaardige opties / pijlers | **Card-rij** | `add_card_row` |
| Losse statistieken | **Big numbers** | `add_big_numbers` |
| Genummerde lijst mét toelichting | **Labeled rows** | `add_labeled_row` |
| Echt een opsomming (en niets beters past) | **Bullets** | layout 1, body-placeholder |
| Afsluiten met een actie | **Statement + CTA** | layout 2, één zin + vervolgstap |

> Staat je inhoud niet in deze tabel? Dan is het waarschijnlijk een combinatie.
> Splits 'm over twee slides: één idee per slide.

---

## "Ik wilde bullets maken" en wat je dan doet

Dit is de reflex die je moet ombuigen. Vier veelvoorkomende gevallen:

| Je hebt 3 bullets... | ...maak er dit van |
|----------------------|--------------------|
| die gelijkwaardige opties/pijlers zijn | **card-rij** (`add_card_row`), 3 cards naast elkaar |
| die opeenvolgende stappen zijn | **stappenflow** (`add_step_flow`), badges + pijlen |
| die een tegenstelling tonen (2 punten) | **vergelijking** (`add_comparison`) |
| die getallen/percentages zijn | **big numbers** (`add_big_numbers`) |
| die echt losse, ongelijksoortige feiten zijn | **dan mag het een bullet-body**, maar hooguit 3-4, kort |

---

## Werkende code per patroon

De colors, helpers en imports staan in [design-components.md](design-components.md).
De placeholder-indices per layout in [han-template-layouts.md](han-template-layouts.md).
Hieronder de twee patronen die vaak vergeten worden, met complete code.

### Quote-slide (layout 5 "Vergelijking")

Deze layout is in de template gebouwd als quote-slide: een zwart vlak rechts met
een aanhalingsteken, witte quote-tekst, en een naambalk eronder. Je hoeft het
vlak niet te bouwen, alleen de placeholders te vullen.

```python
# layout 5 = Vergelijking / Quote
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.placeholders[0].text  = "AI is geen tool die je erbij pakt. Het is een manier van werken."
slide.placeholders[11].text = "Guus Witjes, Team Lead Digital Marketing"
add_notes(slide, "Laat deze even hangen. Stilte na een quote is sterker dan doorpraten.")
```

Gebruik dit voor: een prikkelende stelling, een citaat van een expert/student,
of een kernboodschap die moet blijven hangen. Eén per deck is genoeg.

### Afbeelding-slide (layout 3 "Titel, Tekst en Afbeelding")

Tekst links (idx 12), foto rechts (PICTURE-placeholder idx 11). Gebruik
`insert_picture` op de placeholder, niet `add_picture` met losse coördinaten,
dan valt de foto netjes in het kader.

```python
# layout 3 = Titel, Tekst en Afbeelding
slide = prs.slides.add_slide(prs.slide_layouts[3])
slide.placeholders[0].text  = "Studenten bouwen hun eigen AI-werkwijze"
slide.placeholders[12].text = "Geen losse trucjes, maar een systeem dat met ze meegroeit."
slide.placeholders[11].insert_picture(
    "assets/photos/classroom.png"  # webp eerst converteren, zie brand-assets.md
)
add_notes(slide, "Vertel hier het concrete verhaal achter de foto.")
```

> **Let op:** de HAN-foto's zijn `.webp`. python-pptx slikt webp niet. Converteer
> eerst naar PNG via de helper in [brand-assets.md](brand-assets.md).

### Statement-slide (layout 2, half width)

Eén zin, groot, links. Rechterhelft bewust leeg is rust. Geen body-bullets.

```python
slide = prs.slides.add_slide(prs.slide_layouts[2])
slide.placeholders[0].text = "Wat als elk teamlid een eigen AI-assistent had?"
# body (idx 11) leeg laten of één korte regel
add_notes(slide, "Stel de vraag, wacht, en ga dan pas door.")
```

---

## Het ritme van een goed deck

Een deck dat boeit, wisselt patronen af. Een typisch ritme:

```
Titeldia, Statement, Afbeelding, Card-rij, Bullets, Vergelijking,
Afbeelding, Big numbers, Quote, Statement+CTA
```

Niet:

```
Titeldia, Bullets, Bullets, Bullets, Bullets, Bullets, Bedankt
```

Vuistregel uit [visual-selection-guide.md](visual-selection-guide.md):
**na maximaal twee tekst-slides een visuele slide.**

---

## Checklist voor je opslaat

- [ ] Opent met Titeldia (layout 0), niet met een lijst
- [ ] Minder dan 1/3 van de contentslides is een kale bullet-body
- [ ] Geen twee bullet-slides achter elkaar
- [ ] Minstens één quote/statement én één afbeelding/visueel patroon
- [ ] Elke slide max 6 regels, idealiter 1-3
- [ ] Patronen wisselen elkaar af (geen 5x dezelfde opbouw)
- [ ] Sluit af met een actie, niet met "Bedankt voor de aandacht"
