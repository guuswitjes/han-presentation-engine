# Presentation Engine (HAN) , deelbare skill

Een drop-in Claude Code skill die complete, on-brand **HAN PowerPoint-presentaties**
bouwt met python-pptx. Geen losse bullet-decks, maar afwisselende slide-patronen
(quote, afbeelding, vergelijking, stappen, cards) in de HAN-huisstijl: magenta
(#e50056), Roboto Condensed, HAN-iconen en de schuine stripe.

Deze map is **op zichzelf compleet**: template, iconen, logo's, foto's en fonts zitten
erin. Kopieer hem naar je eigen `.claude/skills/` en het werkt.

---

## Wat zit erin

```
presentation-engine/
├── SKILL.md                  de skill zelf (workflow van doel naar deck)
├── knowledge/                9 naslagdocumenten
│   ├── slide-patterns.md       patroon per slide (anti-bullet), kern van de skill
│   ├── han-template-layouts.md layouts + placeholder-indices
│   ├── design-components.md    python-pptx bouwblokken (cards, stappen, badges)
│   ├── brand-assets.md         iconen, logo, stripe, foto's, webp-conversie
│   ├── worked-example.md       compleet voorbeeld-deck als blauwdruk
│   ├── presentation-structures.md / slide-design-principles.md
│   ├── visual-selection-guide.md / speaker-notes-guide.md
├── scripts/
│   └── embed_fonts.ps1       fonts insluiten voor delen (PowerPoint COM)
└── assets/
    ├── template/han-powerpoint-template.pptx   16:9 HAN-thema, 6 layouts
    ├── brand/                  icons (47), icons-black (12), logos, photos, backgrounds
    └── fonts/                  Roboto + Roboto Condensed (variable, .ttf)
```

---

## Installeren

1. Kopieer de map `presentation-engine/` naar je `.claude/skills/`:
   - Globaal (alle projecten): `~/.claude/skills/presentation-engine/`
   - Of per project: `<project>/.claude/skills/presentation-engine/`
2. Herstart Claude Code of laad de skills opnieuw.
3. Roep aan met `/presentation-engine` of vraag gewoon om een HAN-presentatie.

> De map moet de inhoud direct bevatten (dus `.claude/skills/presentation-engine/SKILL.md`,
> niet een extra tussenmap). `SKILL.md` met hoofdletters: op macOS/Linux telt dat.

### Pad aanpassen

De voorbeelden in `knowledge/` gebruiken:
```python
SKILL = os.path.expanduser("~/.claude/skills/presentation-engine")
```
Installeer je per project in plaats van globaal? Pas dat pad één keer aan.

---

## Vereisten

| Nodig | Waarvoor | Installeren |
|-------|----------|-------------|
| **Python 3** + `python-pptx` | de deck bouwen | `pip install python-pptx` |
| **Pillow** (met webp) | iconen/foto's (webp) naar png | `pip install pillow` |
| **Roboto + Roboto Condensed** | juiste weergave van de deck | installeer de `.ttf`'s uit `assets/fonts/` |
| **PowerPoint** (optioneel) | fonts insluiten bij delen | alleen voor `scripts/embed_fonts.ps1` (Windows) |

**Fonts installeren:** open de vier bestanden in `assets/fonts/` en klik "Installeren".
Zonder deze fonts opent de deck met een vervangend lettertype en klopt de huisstijl niet.

---

## Gebruiken

Vraag bijvoorbeeld:
- "Maak een presentatie van 15 minuten over X voor mijn team"
- "Bouw een deck op basis van dit document"

De skill loodst je langs: doel en duur, een **slide-plan met een patroon per slide**
(zo voorkom je een deck vol bullets), content, speaker notes, en daarna het bouwen.
Zie `knowledge/worked-example.md` voor een compleet voorbeeld dat je kunt nadraaien.

### Fonts insluiten (bij delen)

De deck verwijst naar Roboto bij naam. Deel je 'm met iemand zonder dat font,
dan valt PowerPoint terug op een ander lettertype. Sluit de fonts dan in:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/embed_fonts.ps1 -Path "pad/deck.pptx"
```
(Windows + PowerPoint nodig. Roboto is Apache-2.0 en mag ingesloten worden.)

---

## Licentie en gebruik

Dit pakket bevat **interne HAN-merkassets** (logo's, foto's met herkenbare personen,
iconen, huisstijl-template). **Alleen delen binnen HAN-context, niet publiek
verspreiden.** Zet dit niet in een openbare repository.

- **Roboto / Roboto Condensed:** Apache-2.0, vrij te gebruiken en in te sluiten.
- **Avenir Next** (het web-font van han.nl) zit hier bewust niet in; het is apart
  gelicentieerd. De PowerPoint-huisstijl gebruikt sowieso Roboto Condensed.
- **Foto's:** controleer per foto of je toestemming hebt voor het gebruik buiten de
  oorspronkelijke context.

---

## Onderhoud

Bron van waarheid leeft in de workspace van de maker:
- Skill-logica: `03-han/.claude/skills/presentation-engine/`
- Template: `03-han/_knowledge-base/templates/han-powerpoint-template.pptx`
- Brand-assets: `03-han/_context/brand/han-design-system/`

Werk eerst die bronnen bij, exporteer daarna opnieuw naar deze map. Dit is een
**export voor delen**, geen werkkopie.
