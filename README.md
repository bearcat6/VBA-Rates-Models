# VBA Rates Models

JPY rates analytics prototypes implemented in Excel/VBA.

This repository is organized as a modular VBA library for building JPY OIS discount curves, valuing OIS swaps, fitting SABR smiles, calibrating and simulating a 1-factor Hull-White model, and extending toward CMS spread valuation.

## Scope

The repository is managed by purpose rather than by a flat list of VBA classes.

| Area | Purpose | Status |
|---|---|---|
| `01_discount_curve` | Basic JPY TONA/OIS discount curve construction | Active |
| `02_ois_swap` | OIS swap cashflow and valuation logic | Active / expanding |
| `03_sabr` | Normal SABR smile, density and quantile logic | Planned / expanding |
| `04_hull_white_1f` | Hull-White 1F calibration and Monte Carlo curve simulation | Active |
| `05_cms_spread` | CMS spread valuation; currently mainly swaption volatility handling | Planned / partial |

## Target Repository Structure

```text
VBA-Rates-Models/
├─ docs/
│  ├─ 00_overview/
│  │  ├─ Module_Map.md
│  │  └─ Roadmap.md
│  ├─ 01_discount_curve/
│  │  └─ Design.md
│  ├─ 02_ois_swap/
│  │  └─ Design.md
│  ├─ 03_sabr/
│  │  └─ Design.md
│  ├─ 04_hull_white_1f/
│  │  └─ Design.md
│  └─ 05_cms_spread/
│     └─ Design.md
├─ src/
│  ├─ common/
│  ├─ 01_discount_curve/
│  ├─ 02_ois_swap/
│  ├─ 03_sabr/
│  ├─ 04_hull_white_1f/
│  └─ 05_cms_spread/
├─ examples/
├─ tests/
└─ README.md
```

## Dependency Direction

The intended dependency direction is one-way.

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

`03_sabr` is kept independent from CMS because the same smile, density and quantile logic can also be used for risk analysis and stress scenario construction.

## Current Migration Policy

Existing VBA files are retained until they are safely moved with their full contents preserved. The final module structure and documentation are prepared first so that future class/module moves can be done mechanically and reviewed one file at a time.

See `docs/00_overview/Module_Map.md` for the migration map from the old flat structure to the new purpose-based structure.

## Design Principles

- Keep common utilities separate from product and model logic.
- Keep discount curve construction independent from trade valuation.
- Keep model calibration and simulation independent from Excel sheet I/O.
- Use Excel/VBA-friendly interfaces rather than excessive abstraction.
- Avoid proprietary, customer, or confidential data.

## Disclaimer

This repository is for education, validation and prototyping only. Practical use requires separate checks for model validity, numerical accuracy, boundary conditions, auditability and data governance.
