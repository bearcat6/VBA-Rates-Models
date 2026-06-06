# Design

This document is retained as a top-level pointer for compatibility with older links.

The repository design has moved to a purpose-based modular structure.

## Current Design Documents

| Document | Purpose |
|---|---|
| `docs/00_overview/Module_Map.md` | Overall module map and migration map |
| `docs/00_overview/Roadmap.md` | Development and migration roadmap |
| `docs/01_discount_curve/Design.md` | Discount curve design |
| `docs/02_ois_swap/Design.md` | OIS swap valuation design |
| `docs/03_sabr/Design.md` | Normal SABR design |
| `docs/04_hull_white_1f/Design.md` | Hull-White 1F design |
| `docs/05_cms_spread/Design.md` | CMS spread valuation design |

## Design Direction

The repository is organized by analytical purpose:

```text
common
01_discount_curve
02_ois_swap
03_sabr
04_hull_white_1f
05_cms_spread
```

The old flat structure under `src/classes` and `src/modules` is being migrated toward the purpose-based structure described in `docs/00_overview/Module_Map.md`.

## Dependency Direction

```text
common
  ↓
01_discount_curve
  ↓
02_ois_swap

common + 01_discount_curve
  ↓
04_hull_white_1f

common + 01_discount_curve + 03_sabr
  ↓
05_cms_spread
```

Avoid reverse dependencies and avoid mixing product valuation logic into curve construction classes.
