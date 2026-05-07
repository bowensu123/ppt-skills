# Design Principles for Path A Polish

This is the **design goal** the agent uses when judging whether a slide
is well-composed. The agent reads this document, looks at the rendered
slide, and applies fixes via mutate ops based on its own judgment. There
are NO hardcoded "should be X%" thresholds in the rules — the rules
describe the slide's current state; the agent decides what to change.

---

## The one-line goal

> 信息密集，但通过清晰主结论、明确分区、严格对齐和视觉层级，让观众一眼知道
> 先看什么、重点是什么、结论是什么。

**Translated**: information-dense, but through a **clear main conclusion**,
**distinct partitions**, **strict alignment**, and **visual hierarchy**, the
viewer knows at a glance what to look at first, what's important, and what
the conclusion is.

---

## The five qualities, in priority order

### 1. Clear main conclusion (主结论清晰)

A slide should have **one** dominant element that announces "this is the
takeaway". This is usually the title or a hero number/quote. Everything
else supports it.

**The agent looks for**: is there a single most-emphasized element?
What carries the most visual weight (size × contrast × position)? If
the heaviest element is decorative chrome rather than the title, it's
wrong.

**Failure modes**:
- Title same weight as body text → no hierarchy
- Decorative bars/badges out-shouting the actual content
- Multiple competing focal points

### 2. Distinct partitions (明确分区)

Different content roles should sit in distinct visual zones.
Title-zone, content-zone, footer-zone are clearly separated. Cards
within content-zone are clearly separated from each other.

**The agent looks for**: can you trace the boundary between
sections? Do gaps between cards/sections read as INTENTIONAL whitespace
(>15% of card dimension) rather than accidental drift?

**Failure modes**:
- Cards almost touching, no breathing room between them
- Title overlapping into content area
- Footer crammed against bottom card
- Two zones with no visual separation (line / fill change / spacing)

### 3. Strict alignment (严格对齐)

Edges line up. If two cards are in the same row, their tops AND
bottoms are aligned. If text shapes are stacked, their lefts are
aligned. Off-by-a-few-pixels alignment looks worse than deliberately
asymmetric layout.

**The agent looks for**: imaginary grid lines through the slide. Do
shape edges land on those lines? Are there exactly 1-2 column lines
and 2-4 row lines, or is everything floating freely?

**Failure modes**:
- Card #3's top is 50K EMU below cards #1/#2's tops
- Body text in card #1 starts 10K left of body text in card #2
- Footer text not horizontally aligned with title

### 4. Visual hierarchy (视觉层级)

Order of importance ⇔ order of visual weight. Title biggest. Section
headers next. Body text smaller. Captions smallest. Each level is
distinguished by AT LEAST one of: size, weight, color, position.

**The agent looks for**: count the distinct font-pt values. Should
be 3-5 levels. If only 1-2 levels, no hierarchy. If 7+, chaos. Are
heavier weights used for more-important content?

**Failure modes**:
- Title is 14pt, body is 14pt — no scale separation
- All text is same color (no emphasis on key terms)
- Bold used randomly, not for hierarchy
- Decorative element heavier than content

### 5. Information density (信息密度)

The slide should feel SUBSTANTIAL — content fills the canvas without
crowding. Empty slide = waste. Cluttered slide = unparseable.

**The agent looks for**: content takes 40-65% of slide area
(visually, not sum-of-bbox). Each card is 50-75% filled with
content. White space is generous-but-controlled, not vast empty
zones.

**Failure modes**:
- 4 cards, each 80% empty (info-sparse)
- Content squished against edges (info-overcrowded)
- Single card on a slide with > 60% empty space
- Decorations multiplying to fill empty space rather than real content

---

## How the agent uses this document

### Step 1 — Read the slide composition descriptor

`state_summary.py` produces a `composition.json` per slide that lists
every element's:
- Geometry (position + size + size-relative-to-card)
- Visual weight estimate (area × contrast × font_pt)
- Type / role hints (icon / title / body / decoration)
- Alignment relationships (which other shapes share its top/left/right edge)

### Step 2 — Compare against principles

For each of the 5 qualities, the agent makes a JUDGMENT (using its
multimodal vision):
- Does THIS slide have a clear main conclusion?
- Are partitions distinct enough?
- Is alignment strict enough?
- Is the hierarchy clear?
- Is the density right (not too sparse / not too crowded)?

The judgment is **not** "icon area must be 8%". It's "looking at this
slide, does the icon read as decoration or does it dominate when it
shouldn't?"

### Step 3 — Apply fixes via mutate ops

The agent picks the highest-impact violation and applies one mutate op
to fix it. Examples:

| Violation | Fix |
|---|---|
| Title not heaviest | `set-font-size` title up, or `set-font-bold` |
| Cards drift apart | `equalize-gaps` |
| Card too sparse | `set-font-size` icon/text up, OR `resize` card down |
| Card too crowded | resize card up, OR shrink decorations |
| Mismatched alignments | `align --edge top` |
| Hierarchy flat | `apply-typography --role title/body/caption` per shape |
| No focal point | `set-fill` accent color on hero shape |

### Step 4 — Re-render and re-judge

After each fix, agent reads the new render, asks the same 5
questions. Iterates until "the slide passes the first-glance test"
or further changes risk over-styling.

---

## Anti-patterns the agent should specifically reject

- "Make every element exactly equal" → defeats hierarchy
- "Center everything" → loses tension/direction
- "Use 5 different accent colors" → defeats focal point
- "Fill every empty pixel" → defeats density / breathing room
- "Match the original deck pixel-for-pixel" → if the original is
  poorly designed, perpetuating its choices is wrong

---

## Why no hardcoded thresholds

Earlier versions of this skill had rules like
"icon area should be 8% of card". This produced rigid output that
felt mechanical — same emoji size on every deck regardless of
content density. The principles above let the agent calibrate per
deck:

- A poster slide with one big takeaway → larger icon is right
- A dashboard slide with 6 dense cards → smaller icon is right
- A kid-friendly explanatory slide → playful larger emoji is right

The agent reads the slide's content, judges what design decisions
serve THIS specific slide's communication goal, and applies them.
