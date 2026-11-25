# REVEAL 5D — Context (read this first, nothing else unless needed)

Reference repo for the REVEAL Residential Compound 3D / 4D / 5D page (model + schedule + cost), Al-Obour.
Owner: Saeed. The kit itself is generic (`README_KIT.md`); this file covers how it is used for REVEAL.
Geometry comes from [reveal-4d](https://github.com/REVEAL-Compound/reveal-4d); this repo does not keep its own drawings.

## Build (one command)
```
examples/reveal-all/build.sh <reveal-4d clone> <baseline master.xlsx> <actualized master.xlsx> <data date> [out]
# current issue:
examples/reveal-all/build.sh ../reveal-4d master_v03.xlsx master_v04.xlsx 2026-10-03
```
Output: `out/reveal-all/index.html` (3D + 4D + 5D tabs, ~30 MB, one file) + Programme / Update / Budget_Baseline /
Cost_Report workbooks + XER. The page is sent as a file (not published as an artifact: project data).

## Sources
| Input | Where | Rule |
|---|---|---|
| Geometry | reveal-4d `data/geometry/*.json`, `SITE.json`, `landscape.json`, `terrain.json` | read only |
| Baseline | -REVEAL-Schedule-Control tag `master-v03` | sheet `Master Schedule موحد` |
| Update (actuals) | -REVEAL-Schedule-Control tag `master-v04` (actualized from IR/CPR, data date 2026-10-03) | new tag = new update |
| Budget | Master cost loading (BOQ rates incl. 25% margin) — Saeed's choice | not the Cost Control (cost-price) model |

Data files made by the adapter (`data.json`, `activities.json`, `links.json`, `*_Update.xlsx`) are git-ignored.

## Decisions (Saeed)
- Budget basis: Master rates (2026-11-22). Actual cost (AC) not linked yet: AC = EV. Linking AC from
  REVEAL-Cost-Control was put on hold (2026-11-23) — cost basis differs (cost price vs selling), see the note below.
- Whole project in one page: WB, EB, NB, SB + Mall (2026-11-22).
- Model moves daily; every figure on the page is monthly (cards, month table, month activity list) (2026-11-24).
- Only the actual (update) model is drawn; baseline figures still drive the planned numbers (2026-11-24).
- Key cards float on the view as glass cards in the free screen space; settings stay in the side panel; colour legend removed (2026-11-23).
- Site works and landscape start after the last concrete activity on the whole site (2026-11-23): `--site-start global`.
- Pump stands: from the pump standing study, `examples/reveal-all/pump_stands.json` (2026-11-23).
  EB from the external east street; west/east internal roads too narrow (8.6 m road, 11.5 m between basements);
  middle road usable with conditions. NB1-NB3 have no stand Dec-2027 to Jun-2028 — decision open.

## Assumptions (ASSUMED / ESTIMATE, shown as such on the page)
- Excavation pit per schedule group = basement footprint + 1.0 m work space; bottom = footing underside − 0.10 PCC − 1.5 m soil replacement; vertical sides.
- No-standing strip around an open pit = pit depth (1:1 from the bottom). Stand after backfill: 2 m off the basement wall.
- Pour day = last day of each concrete activity. In-progress % at the data date = elapsed working days.
- Site works stages: infrastructure + kerbs 0–60 d, paving + pools 45–120 d, softscape 100–180 d (PROPOSED activities `SW-*`, not in the Master, no budget).
- Finishing colours by activity name: block work = raw, بياض = plaster grey, رخام/جرانيت/سيراميك = stone, بازلت = basalt, دهان / مونة ملونة = final colour; finishing under way = blue.
- Retaining wall 0.3 m on the basement perimeter and soil layer 0.5 m under footings where the drawings carry none.

## Known gaps
- Labour: zero (Master has no crews).
- Clubhouse: not on the site plan (MASTER1.dxf checked 2026-11-23) and no Master activities.
- Mall Z3 retaining wall / waterproofing / external backfill (~2.5 M) have no geometry; WB excavation is not in the Master.
- No CPM logic in the Master (no critical path / float).

## Cost basis note (for linking AC later)
Cost Control (WB v28 / EB v03 / PRJ v01, data date 2026-10-04) is at cost price; the 5D is at Master (selling) rates.
Options offered: scale AC by each block's Cost-Control CPI, use AC as is, or move the 5D to cost price. Waiting for Saeed.
