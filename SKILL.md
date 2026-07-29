---
name: sd-worx-brandbook
description: Apply the official SD Worx brand guidelines to any visual or written deliverable. Use this skill whenever you are producing something that carries the SD Worx name or will be seen by SD Worx colleagues, customers, partners or management, even when the user does not mention branding. That includes slide decks and presentations, HTML tools and dashboards, one-pagers, reports, diagrams, charts, mockups, internal MT material, and LinkedIn or Viva Engage posts. Trigger on explicit requests such as "follow the SD Worx brandbook", "make this on-brand", or "SD Worx style", and equally on implicit ones such as "build me a deck for the MT", "make an HTML tool for the sales team", or any request that names an SD Worx product, team or customer. If in doubt about whether a deliverable is SD Worx branded, apply this skill.
---

# SD Worx Brandbook

Produce work that looks like SD Worx made it, first time, without a styling round.

The brand rests on one idea: **we turn complexity into confidence**. The visual system says the same thing. Blue dominates, white gives room to breathe, accents appear rarely and mean something when they do. Restraint is the point. If a layout looks busy or colourful, it is off-brand no matter how well each individual rule was followed.

## Non-negotiables

Check these before delivering anything. Most brand failures are one of these five.

1. **Blue leads.** SD Worx Blue `#006dd8` and white carry the layout. Everything else supports.
2. **Accents stay under 10%.** Red `#f1002f` or yellow `#ffbe00`, one of them per layout, never both, never as a full-bleed background.
3. **Text is black, white, or SD Worx Blue.** Never a light or accent colour. Never a colour that drops below WCAG AA.
4. **Black never becomes a background.** It is for text, icons and small details only. For dark surfaces use Deep Navy `#001c52` or Persian Blue `#000d3a`.
5. **Space is a feature.** Generous margins, consistent padding, few elements. When unsure, remove something.

## Colours

### Primary

| Name | Hex | Role |
|---|---|---|
| SD Worx Blue | `#006dd8` | Dominant. Backgrounds, gradients, key UI, headings, emphasis |
| White | `#ffffff` | Second dominant. Backgrounds, white space, text on blue, widget surfaces |
| Black | `#000000` | Text, icons, small graphic detail, widget borders. Never a background |

### Accents, used sparingly

| Name | Hex | Role |
|---|---|---|
| SD Worx Red | `#f1002f` | Calls to attention, CTA highlight, single emphasis point |
| SD Worx Yellow | `#ffbe00` | Calls to attention, CTA highlight, single emphasis point |

One accent per layout. Under 10% of the surface. Never a full-bleed red or yellow background.

### Secondary blues

Use to differentiate within charts, infographics, UI widgets and background variation. They expand the palette without introducing a new hue.

| Name | Hex |
|---|---|
| Icy Blue | `#d9f1ff` |
| Cool Sky | `#9ed2ff` |
| Brilliant Azure | `#38a1ff` |
| Steel Azure | `#006dd8` |
| Deep Navy | `#001c52` |
| Persian Blue | `#000d3a` |

### Verified contrast pairs

Reach for these rather than inventing combinations. All pass WCAG AA.

| Pair | Ratio |
|---|---|
| White on `#006dd8` (and reverse) | 5.02:1 |
| Black on white (and reverse) | 21:1 |
| `#ffbe00` on `#006dd8` (and reverse) | 12.61:1 |
| White on `#000d3a` | 18.77:1 |
| White on `#001c52` | 16.3:1 |
| Black on `#38a1ff` | 7.71:1 |
| Black on `#9ed2ff` | 13.11:1 |
| Black on `#d9f1ff` | 17.97:1 |
| Black on `#f1002f` | 4.78:1, large text and graphics only |

Note that white on `#006dd8` passes AAA only at large sizes. For long body copy on blue, prefer black text on a white or Icy Blue surface instead.

### Charts

Build series from the blue ramp: Deep Navy, Steel Azure, Brilliant Azure, Cool Sky, Icy Blue. Use one accent only to mark the single data point that carries the message. Do not colour every series differently for decoration.

## Typography

Titles use **SD Worx Display**. Body, values, labels and anything below 14pt use **Inter**.

In HTML, use Inter throughout. SD Worx Display is licensed and will not resolve in a browser, and the brandbook already sanctions Inter at smaller sizes. Load Inter from Google Fonts and set a sensible stack:

```css
font-family: 'Inter', -apple-system, 'Segoe UI', system-ui, sans-serif;
```

In PowerPoint and Word, set SD Worx Display on titles and Inter on body. On a machine with the corporate fonts installed it resolves correctly; elsewhere it degrades to Inter, which is the sanctioned fallback anyway.

Keep the hierarchy shallow. A title, a subtitle, body, and a small label is usually enough. Never fewer than three visible steps of size difference between title and body, or the hierarchy reads as flat.

> Gap: the official typographic scale, weights and line heights were not available when this skill was written. Sizes below are a reasonable default, not brandbook text. Replace them if the typography page becomes available.

Default web scale: 48/32/24 for display and headings, 16 for body, 13 for labels, 1.5 line height on body, 1.2 on headings.

## Layout and grid

**Margin.** Divide the short side of the format by 16. That is the margin, and it is also the base unit **X**.

**Grid.** For wide formats (a 16:9 slide, a desktop screen), divide the short side into 8 and the long side into 16. For narrow formats, divide both sides into 8.

**Spacing.** Size and space everything in multiples or clean fractions of X: 0.25X, 0.5X, 1X, 1.5X, 2X, 3X. This is what makes layouts feel structured rather than assembled.

### The four layout options

Pick by how permanent the touchpoint is and how much there is to say.

| Layout | Use when |
|---|---|
| Pillar with gradient | Key or permanent touchpoint, lots of content, message needs clarity |
| Pillar with image | Key or permanent touchpoint, lots of content, message needs human connection |
| Gradient, full surface | Brief or one-off, message short and sharp |
| Image, full surface | Brief or one-off, message engaging and authentic |

Vary across a deck or a document. Using one layout for every slide is a listed anti-pattern.

### Pillar layout

The three pillars in the logo represent employers, workers and regulators, forces that never quite align. The layout system extends those pillars as containers. They can be rotated, moved, resized, split or merged, but they always align to the grid.

- One or two clearly defined content areas, three containers maximum.
- White containers hold text. They pair with a container holding either an image or a gradient.
- Never put a gradient and an image in the same container.
- Keep strong contrast between positive and negative space. Bold and simple beats intricate.

## Widgets

Small cards that surface one key number or message. They are the workhorse of dashboards, HTML tools and data slides.

**Structure**
- White background, always. Widgets are the calm pause inside a composition.
- 1px black border. No shadows, no glows, no dimensional effects.
- Border radius 8px for containers, cards and CTAs. 4px for pills, status labels and small elements. Digital follows a strict 8-point system.
- Padding 16 to 24px on all sides. 12 to 16px between content groups.

**Content hierarchy**, in this order:
1. Title, short and functional, SD Worx Display (Inter below 14pt)
2. Primary value or key message, the visual focal point, Inter
3. Supporting detail, one or two lines maximum, Inter
4. Optional status label such as APPROVED, PENDING, REFUSED, Inter

Text is black by default, SD Worx Blue for secondary emphasis. Status colours sparingly.

If a widget needs a paragraph, it is the wrong component.

## Logo

Three versions:
- **Primary**, the horizontal lockup. Default choice everywhere.
- **Secondary**, the stacked lockup. Only when horizontal space genuinely does not allow the primary.
- **Brand mark**, the pillars alone. Avatars, favicons, app icons. Never in marketing material where the full logo fits.

Rules that get broken most:
- Clear space of 1/2 X on all sides, where X is half the height of the brand mark. Nothing enters it.
- Minimum digital height 24px, print 10mm. The stacked version needs 50px digital, 20mm print.
- Use the full-colour logo whenever the background allows. Switch to monochrome white or black only when the background interferes.
- Never distort, recolour, rotate, or place on a busy background.

Assets are in `assets/`. See `references/assets.md` before using them.

## Voice

Personality: hands-on, confident, human, grounded. In practice that means short sentences, concrete nouns, and no hedging. Take a position. Explain the hard thing plainly rather than dressing it up.

Say "we turn complexity into confidence" and "we make work work" only where they earn their place. Repeating taglines inside internal material reads as filler.

Proof points, when scale needs evidence: 100K+ employers supported, 6M+ employees paid monthly, presence in 30 European markets, 80 years of experience, 26 own payroll and time engines, 10,000+ HR experts.

The four value words are compliance, clarity, continuity, confidence. Confidence is the outcome of the other three, so do not list it as a peer.

## Producing HTML

Read `references/html.md` for the ready-made token block and component patterns. Copy the tokens rather than retyping hex values, which is where drift starts.

## Producing slides

Read `references/pptx.md` for slide dimensions, the grid translated to a 16:9 canvas, and the master layouts.

## Before you deliver

Run this check. State the result briefly rather than silently assuming it passed.

- Is blue clearly dominant, with white second?
- Is there at most one accent colour, under 10% of the surface?
- Is all text black, white, or SD Worx Blue, and does every pair pass AA?
- Is black used only for text, icons and detail, never as a background?
- Do widget borders, radii and padding match the 8-point system?
- Does spacing derive from the base unit X?
- Does the logo have its clear space and meet minimum size?
- Is there anything on the surface that could be removed without losing meaning?

## Known gaps in this skill

Be honest about these rather than inventing an answer. If the user asks about one, say it is not covered and propose something explicitly as a suggestion.

- **Typographic scale.** Not available. The sizes above are a default, not brand law.
- **Gradients.** The brandbook has dedicated gradient and gradient-on-typography sections that were not available. The working rule here is that gradients run between colours in the sanctioned blue palette only, never crossing into red or yellow. That is an inference from the colour principles, not a quoted rule.
- **Photography and AI photography.** Not covered. The brandbook specifies a signature style and a product style on light grey or blue backgrounds, but the detail is missing. Do not generate or specify imagery confidently.
- **Graphical elements** and **Brand Motion.** Not covered.
- **Official MS Office templates.** The brandbook lists them under Touchpoints. If the user has the .potx, it beats anything in this skill.
