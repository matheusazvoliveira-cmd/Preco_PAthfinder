# MRS Pathfinder GDT Cost Proposal

This folder contains the cost-consolidation work for the MRS Pathfinder GDT railcar proposal. The objective is to combine the Pathfinder equipment BOM, SKID additions, Brazilian PAC costs, import-cost assumptions, and the GDT container/railcar integration rollup into a traceable cost model.

## Main Deliverable

`Lista_Final_Pathfinder_SKID_Custos_English_Updated.xlsx` is the English workbook prepared for review and submission. It contains:

- `Executive Summary (GDT Project)`: project-level totals linked by formulas to the detailed sheets.
- `Final Cost List`: consolidated Pathfinder and SKID items with part number, description, system, subsystem, quantity, direct/nationalized USD cost, gross USD cost including taxes, BRL totals, source, and status.
- `GDT Integration and Assembly`: container adaptation, integration materials, labor, engineering, transportation, training, commissioning, warranty, and SG&A from the GDT cost rollup.
- `Missing Items`: items still requiring a price or confirmation.

The original-language workbook and the latest consolidated workbook are also retained as source/reference files.

## Source Files

- `Pathfinder BOM.xlsx`: original Pathfinder equipment BOM.
- `SKID BOM.xlsx`: SKID BOM. SKID part numbers already present and priced in the Pathfinder BOM are not counted again.
- `7-PAC Item Costs.xlsx`: Brazilian PAC item-cost database. PAC is treated as the local Brazilian cost reference; import TP, international freight, and import duty are not applied again to PAC items.
- `Pathfinder GDT MRS - Cost Rollup - Rev0.xlsx`: container, railcar adaptation, integration, labor, engineering, and field-service rollup.
- `4 - Cotação - Planilha Modelo - Vagao.xlsm`: original quotation model used to understand the import and tax calculation structure.

## Cost Rules

### PAC items
When a part has a non-zero PAC value, the PAC BRL value is used as the final local Brazilian cost reference. Taxes already included in PAC are not grossed up again.

### Imported BOM items
For items without PAC but with a US BOM price, the model applies:

1. Transfer Price: 17.2%, unless the item is marked green in the color-coded review copy.
2. International freight and logistics: 9%.
3. Import duty (II): 16%.
4. Domestic freight: 3%.
5. Output-tax gross-up for the gross-price view: PIS/COFINS 9.25% and ICMS 12%.

No profit margin is included. The gross-price view is a tax-recomposed breakeven view, not a selling price with profit.

### Color-coded review rules
The color copy uses these rules:

- Yellow: Transfer Price remains included.
- Green: direct supplier purchase; Transfer Price is excluded, while freight and applicable taxes remain.
- Red: remove the item entirely from the final list and totals.
- Blue: item is made by GET; retain it for GET impact analysis and mark the source as GET.

The international freight assumption was changed from 15% to 9% for the updated workbook.

## Important Item Decisions

- Red items were removed from the final list and totals.
- `17FL519A1` was replaced by `17FL472A2` for the TMC panel. The replacement is currently marked missing until its price is confirmed.
- `17FB193E2F` was added as a GET card and retained as a GET item pending price confirmation.
- `17FM839B1` is the priced TCD / Thin Client Display reference. Its PAC value is used when available; the SKID variant must not be counted as a second display without confirmation.
- `17FM878A6` quantity is 3, derived from the Pathfinder BOM total amount divided by the unit amount.
- Antenna Kit and PTC Kit sub-items are not counted as additional kit costs when covered by their master kit prices.
- Headlight-related items were excluded according to the review rules.

## Scripts

Use the proposal virtual environment at `C:\venvs\proposal`:

```powershell
C:\venvs\proposal\Scripts\Activate.ps1
```

The comparison scripts can be run from this folder:

```powershell
& "C:\venvs\proposal\Scripts\python.exe" "compare_bom_cost.py"
& "C:\venvs\proposal\Scripts\python.exe" "compare_skid_bom_cost.py"
& "C:\venvs\proposal\Scripts\python.exe" "check_quote_coverage.py"
```

`openpyxl` is required. The quotation workbook may be locked by Excel or OneDrive; scripts that read it use a temporary PowerShell copy to avoid changing the original file.

## Review Notes

- The final workbook is a cost model, not a sales-price model.
- Missing prices must be confirmed before the total is considered complete.
- PAC and US BOM prices are different cost bases. They should not be summed for the same physical item.
- The Executive Summary should be opened in Excel so formulas recalculate against the detailed sheets.
