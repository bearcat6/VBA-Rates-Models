# VBA Rates Models

VBA prototypes for JPY rates analytics, including OIS curve construction,
cap volatility bootstrapping, ATM swaption volatility interpolation, and CMS-related valuation utilities.

## Class naming note

`src/classes/clsATMSwaptionVol.cls` holds an ATM swaption volatility matrix and returns ATM volatilities for arbitrary Expiry × Tenor points.

The class name is intentionally ATM-specific to distinguish it from future non-ATM swaption volatility surfaces or smile models.

This repository is for educational and prototyping purposes only.
All sample data are fictional.
No proprietary or confidential data are included.
