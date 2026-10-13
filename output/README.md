# Sola Kitchen: marketplace delivery board review

| File | What it is |
|---|---|
| `Sola_Kitchen_Board_Recommendation.pdf` | 5-slide board deck |
| `Sola_Kitchen_Financial_Model.xlsx` | Formula-driven model (raw data in the `Data_*` tabs; start at `README` → `Assumptions` → `Exec_Summary`) |
| `QA_Summary.md` | Data audit, checks performed, limitations, open questions |
| `AI_Usage_Note.txt` | AI disclosure |
| `build_scripts/` | Python/Node scripts that regenerate the model, the charts and the PDF from `Sola_Kitchen_Data.xlsx` |

To rebuild, run these from a scratch directory that contains the source data as `in/d.xlsx`:
`python3 build_model.py in/d.xlsx model.xlsx` → recalculate in LibreOffice → `python3 charts.py model.xlsx` → `python3 build_deck.py` → `node render.js deck/deck.html deck/deck.pdf`.
