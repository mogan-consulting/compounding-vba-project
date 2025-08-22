# Compounding VBA Project

Goal: reliable compounding functions (annual/monthly/daily) with a tiny test harness.

## Structure
- `src/` VBA modules (.bas)
- `docs/requirements.md` functional requirements
- `tests/` manual/automated test inputs & saved results
- `results/` generated reports
- `changelog.md` change history

## Run
1) Excel → Alt+F11 → VBA Editor → **File → Import File…** → pick `src/compounding.bas`.
2) Back to Excel → Developer → **Macros** → run `Test_Compounding` to fill the active sheet.
3) (Optional) run `SaveResultsToTests` to export the active sheet as `tests/test_results.xlsx`.
