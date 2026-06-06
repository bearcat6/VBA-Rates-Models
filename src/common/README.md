# common

Common VBA utilities shared by curve, product and model modules.

## Intended Contents

- Date utilities
- Day count functions
- Business day adjustment
- Array utilities
- Numerical methods
- Random number generators
- Normal distribution helpers
- Bachelier / normal option pricing utilities if used across modules

## Current Migration Target

| Current Path | Target Path |
|---|---|
| `src/classes/clsRandomNormal.cls` | `src/common/classes/clsRandomNormal.cls` |

Keep this module independent from all product and model modules.
