# OIS Swap Valuation Design

## Purpose

`02_ois_swap` contains OIS swap trade representation, cashflow generation and valuation logic.

This module should use discount curves but should not be mixed into discount curve construction itself.

## Responsibilities

- Hold OIS swap product terms.
- Generate fixed and floating leg schedules.
- Calculate fixed leg PV and floating leg PV.
- Calculate NPV using payer/receiver sign convention.
- Calculate par rate.
- Export cashflow tables for Excel review.

## Main Classes

| Class | Role |
|---|---|
| `clsOISSwap` | OIS swap product and valuation class |
| `clsOISLeg` | Future candidate for leg-level decomposition |
| `clsCashflow` | Future candidate for cashflow-level representation |

## Sign Convention

Leg-level PVs are positive standalone values.

```text
PAYER    = FloatingLegPV - FixedLegPV
RECEIVER = FixedLegPV - FloatingLegPV
```

Par rate is independent of payer/receiver.

## Dependency Rule

Allowed:

```text
02_ois_swap -> 01_discount_curve
02_ois_swap -> common
```

Not allowed:

```text
02_ois_swap -> 04_hull_white_1f
02_ois_swap -> 05_cms_spread
```

## Notes

If curve bootstrapping temporarily uses OIS swap valuation internally, keep the implementation narrow and avoid making curve classes product-heavy. The long-term direction is to separate curve construction, product cashflows and valuation workflows.
