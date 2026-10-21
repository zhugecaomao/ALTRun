**English** · [中文](Calculator-zh)

# Calculator

Type an expression directly: `12*(3+4)`, `=2^10`; `Enter` copies the result. Supports `+ - * /`, powers `^` (or `**`) and parentheses; results keep at most two decimals (`10/3` shows `3.33`; numbers below 0.005 keep one digit after the first significant digit, for example `0.004`).

With "Add structural results (main bars / rebar area)" (`Features.Calculator.StructuralCalc`) turned on in Preferences → Calculator, two lines of structural results are added below when the result is more than 2 × the cover:
- The result as a **beam width (mm)**: number and spacing of main bars. Number = ⌈(width − 2 × cover) / max spacing⌉ + 1, spacing = (width − 2 × cover) / (number − 1)
- The result as a **rebar area As (mm²)**: bars needed for each size (⌈As / area of one bar⌉, area of one bar = π d² / 4)

For example `300*2` (600) shows `Beam width 600 mm: 3 main bars @ 260 c/c` and `As = 600 mm²: 5H13  3H16  2H20  2H25  1H32` by default. The parameters are on the same page:

| Setting | Key | Default | Description |
|---|---|---|---|
| Rebar cover (mm) | `RebarCover` | `40` | Each side, mm |
| Max spacing (mm) | `MaxBarSpacing` | `300` | mm; more main bars are added when the spacing would exceed it |
| Bar sizes (mm) | `BarSizes` | `[13, 16, 20, 25, 32]` | Diameters (mm) listed in the rebar area line |
| Bar prefix | `BarPrefix` | `H` | Bar designation, for example `T`, `Y`, `Φ` |

![Preferences → Calculator](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/prefs-calculator.png)

### Unit and currency conversion
Syntax: number + unit + `in` / `to` / `=` / `->` / `转` + target unit; units are case-insensitive, and `m²` can be written `m2`. For example `10 km in mi`, `5ft to cm`, `100 f to c`, `3 亩 in m2`, `1 GB to MB`, `300 kN to kip`, `20 MPa in psi`. `Enter` copies the number.

| Category | Units |
|---|---|
| Length | mm cm m km in ft yd mi nmi (Chinese names work too) |
| Mass | mg g kg t lb oz 斤 两 |
| Area | mm2 cm2 m2 km2 ha acre ft2 in2 亩 |
| Volume | ml cl l m3 cm3 ft3 gal qt pt cup floz |
| Speed | m/s km/h mph knot ft/s |
| Time | ms s min h day week year |
| Data | bit B KB MB GB TB (powers of 1024) |
| Pressure / stress | Pa kPa MPa (N/mm2) bar psi ksi atm psf |
| Force | N kN lbf kip kgf tf |
| Energy / power | J kJ cal kcal Wh kWh / W kW hp |
| Temperature | C F K |

**Currency conversion** (`100 usd to sgd`) is off by default: tick "Currency conversion" in Preferences → Calculator (`Features.Calculator.Currency`). Once on, the reference rates published by the European Central Bank and other central banks are downloaded once a day from [Frankfurter](https://frankfurter.dev) (free, no sign-up) and saved in `Data\Currency.json`; apart from GitHub, this is the only website ALTRun contacts on its own (a web search's site is visited only when you click "Download Site Icon"). Results show the date of the rates.
