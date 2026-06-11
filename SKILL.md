---
name: presentation-engine
description: Bouwt complete HAN PowerPoint presentaties met python-pptx op de HAN template. Loodst je van doel via slide-plan naar afgeleverd .pptx, met gevarieerde slide-patronen (quote, afbeelding, vergelijking, stappen) in plaats van bullet-lijsten, HAN-iconen en stripe, en optionele font-insluiting.
user-invocable: true
---

# Presentation Engine (Skill)

**ROL:** Je bouwt complete, mooie HAN PowerPoint presentaties. Van doel tot
afgeleverd .pptx, met afwisselende slide-patronen die het hele HAN design system
benutten. Niet een stapel bullet-lijstjes.

---

## Wanneer deze skill

- Gebruiker vraagt om een presentatie, slide deck of PowerPoint
- Presentatie is voor HAN-context (vergaderingen, trainingen, workshops, teamdagen)

## Wat deze skill NIET doet

- Presentaties voor Step Ahead AI (eigen stijl)
- Content schrijven zonder input van de gebruiker

---

## De gouden regel van deze skill

> **Een goed deck wisselt patronen af. Bullets mogen, alleen niet op elke slide.**

Het meest gemaakte foutje is élke slide vullen met een body-placeholder vol bullets.
Een bullet-slide hier en daar is prima waar dat de beste vorm is; het wordt saai
als het de standaard is. Daarom kies je **per slide bewust een patroon** (quote,
afbeelding, vergelijking, stappen, cards, big numbers, statement, of inderdaad
bullets). Het slide-plan in stap 2 is het mechanisme dat afwisseling bewaakt. De
catalogus en grenzen staan in [knowledge/slide-patterns.md](knowledge/slide-patterns.md).
Lees die file altijd voor je gaat bouwen.

---

## Knowledge-bestanden

| Wat | File |
|-----|------|
| **Patroon kiezen per slide (anti-bullet)** | [knowledge/slide-patterns.md](knowledge/slide-patterns.md) |
| Layouts + placeholder-indices | [knowledge/han-template-layouts.md](knowledge/han-template-layouts.md) |
| Python-componenten (cards, stappen, badges, helpers) | [knowledge/design-components.md](knowledge/design-components.md) |
| Logo, iconen, stripe, foto's, webp-conversie | [knowledge/brand-assets.md](knowledge/brand-assets.md) |
| Compleet voorbeelddeck (blauwdruk) | [knowledge/worked-example.md](knowledge/worked-example.md) |
| Opbouw per duur (5/15/30/60 min, dag) | [knowledge/presentation-structures.md](knowledge/presentation-structures.md) |
| Ontwerpprincipes | [knowledge/slide-design-principles.md](knowledge/slide-design-principles.md) |
| Beeldkeuze per slide | [knowledge/visual-selection-guide.md](knowledge/visual-selection-guide.md) |
| Speaker notes | [knowledge/speaker-notes-guide.md](knowledge/speaker-notes-guide.md) |
| Font insluiten (script) | [scripts/embed_fonts.ps1](scripts/embed_fonts.ps1) |

**Template:** `assets/template/han-powerpoint-template.pptx`
(16:9, HAN-thema E50056 + Roboto Condensed SemiBold/Roboto, 6 layouts).

> **Paden:** alle bestanden (template, assets, fonts, scripts) zitten in deze skill-map
> en worden skill-relatief aangeduid. Prefix ze met het absolute pad naar deze
> skill-map (bijv. `~/.claude/skills/presentation-engine`) bij het bouwen.
> Installatie en vereisten: zie [README.md](README.md).

---

## Workflow: van doel naar deck

Je loodst de gebruiker stap voor stap. Vraag niet alles tegelijk; werk de stappen af
en leg per fase iets voor.

### Stap 1 - Doel, publiek, duur

Vraag (als niet gegeven):
- **Doel:** wat wil je bereiken? (kennis overbrengen, enthousiasmeren, afspraken maken)
- **Publiek:** voor wie? (team, MT, studenten, externen)
- **Duur:** hoe lang? (5 / 15 / 30 / 60 min, hele dag)

Pak de bijpassende opbouw uit [presentation-structures.md](knowledge/presentation-structures.md)
en bepaal het aantal slides.

### Stap 2 - Slide-plan met patroon per slide (KRITIEK)

Maak een tabel. Per slide: titel, kernpunt, **en welk patroon**. Dit is de stap die
het bullet-probleem voorkomt. Gebruik de beslistabel uit
[slide-patterns.md](knowledge/slide-patterns.md).

| # | Titel | Kernpunt | Patroon |
|---|-------|----------|---------|
| 1 | ... | ... | Titeldia |
| 2 | ... | ... | Statement |
| 3 | ... | ... | Afbeelding |
| ... | ... | ... | Card-rij / Stappen / Quote / ... |

Controleer het plan tegen de **harde grenzen** uit slide-patterns.md:
- Opent met Titeldia
- Minder dan 1/3 kale bullet-slides, nooit twee achter elkaar
- Minstens één quote/statement én één afbeelding/visueel patroon
- Sluit af met een actie

Leg dit plan voor aan de gebruiker en pas aan op feedback voor je verder gaat.

### Stap 3 - Content per slide

Vul per slide de tekst in: max 1-3 regels, details gaan naar speaker notes.
Bepaal welke slides een **foto, icoon of stripe** krijgen
(zie [brand-assets.md](knowledge/brand-assets.md) en
[visual-selection-guide.md](knowledge/visual-selection-guide.md)).

### Stap 4 - Speaker notes

Per slide notes volgens [speaker-notes-guide.md](knowledge/speaker-notes-guide.md):
opening, kern, overgang. Nederlands tenzij anders gevraagd. Geen herhaling van de slide-tekst.

### Stap 5 - Bouwen

Schrijf één python-pptx script (blauwdruk: [worked-example.md](knowledge/worked-example.md)):
1. Laad de template, verwijder de 6 voorbeeldslides.
2. Bouw elke slide volgens het gekozen patroon:
   - Eenvoudige tekst/afbeelding: template-placeholders (indices in [han-template-layouts.md](knowledge/han-template-layouts.md)).
   - Quote: layout 5. Statement: layout 2. Afbeelding: layout 3.
   - Visueel rijk (cards, stappen, big numbers): components uit [design-components.md](knowledge/design-components.md).
   - Iconen/stripe/logo: helpers uit [brand-assets.md](knowledge/brand-assets.md). Converteer webp eerst naar png.
3. Voeg speaker notes en afbeeldingen toe.
4. Sla op in de projectmap van de presentatie.

### Stap 6 - Controle

Loop de checklist uit [slide-patterns.md](knowledge/slide-patterns.md) af. Scan op
banned words (banlist). Klopt het ritme (geen reeks identieke slides)?

### Stap 7 - Fonts insluiten (optioneel, bij delen)

Standaard verwijst de deck naar Roboto bij naam. Op een pc zonder Roboto Condensed
valt PowerPoint terug op een ander font en klopt de huisstijl niet. Deel je de deck?
Sluit de fonts in:

```
powershell -ExecutionPolicy Bypass -File scripts/embed_fonts.ps1 -Path "pad/deck.pptx"
```

Sluit af met:
```
Presentatie gebouwd: [pad/bestandsnaam.pptx]
Slides: [n]  |  Patronen: [opsomming, bv. titeldia, statement, afbeelding, cards, quote]
Speaker notes: [ja/nee]  |  Fonts ingesloten: [ja/nee]
```

---

## Guardrails

- **Patroon per slide.** Bullets mogen, niet op elke slide. Handhaaf de grenzen uit slide-patterns.md
- Elke slide max 6 regels (streef 1-3). Details in speaker notes
- Kleur zit in shapes (card-border, badge), niet in losse tekst. Max 2-3 kleuren per slide, rood als accent
- Template-placeholders voor simpele slides, components voor visueel rijke slides
- Iconen: alleen de HAN-set, rood óf zwart per deck consistent. Stripe en logo spaarzaam (één stripe per compositie, geen logo bovenop de master)
- webp-assets eerst naar png converteren (python-pptx slikt geen webp)
- Speaker notes in het Nederlands tenzij anders gevraagd
- Geen website-URLs in slides
- Controleer output op banned words (banlist), geen em-dash
