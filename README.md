# ZSP-Scrapping
DMND0025000-ZSPF-scrapping

Spares scrap upload: report `ZMMSPARESCRAP`, tcode `ZSCRAP` (source: `ZMMSPARESCRAP.abap`, spec: the xlsx in this repo).

## Setup
| Item | Value |
|---|---|
| Number range | object `ZSCRAP`, interval `01`, 0000000-9999999 |
| TVARVC | `ZSCRAP_PLANT`, type S, one row per allowed plant (1MIO, 1MIX - business to confirm) |
| Tcode | `ZSCRAP` -> report `ZMMSPARESCRAP`; authorization to be assigned (Helen Xu) |
| BASIS | SAPoffice/attachment storage for the approval email |

The single file is split by banner comments (`ZMMSPARESCRAP_TOP/_SEL/_CLS/_MAIN`) so it can be turned into real includes.

## Behaviour
- Upload via frontend file dialog, binary read, `CL_FDT_XL_SPREADSHEET` (GUI and WebGUI/Fiori). Columns are read by position (A-L).
- Whole file validated first; any error = nothing posted. Test run (default on) ends with "File can be uploaded. Please untick Test Run".
- Per source stock: Unrestricted 344 -> 555, QI 350 -> 555, Blocked 555. Cost centre 3127, reason 0004, header text `Spare Disposal SP<ZSCRAP no>`.
- If 555 fails after 344/350 posted, the transfer is reversed with 343/349.
- Approval email stored once and linked to every material document (BUS2017, relation ATTA); ALV log and emailed XLSX result.

## Assumptions / open points
1. Stock check uses MCHB only (as in spec) - non-batch-managed materials would fail the check.
2. 555 is posted with GM_CODE 03 (goods issue); the spec says "Goods Receipt".
3. Non-batch-managed materials: batch and valuation type are left blank in the BAPI (spec wording is unclear).
4. Number range interval assumed `01`.
5. Extra validations beyond the spec: UoM must equal base UoM, no negative / zero quantity, non-numeric cells.
6. The Disposal Qty column is a formula in the sample - the file must be saved with calculated values.
7. A row with several stock types can be partly posted (e.g. Unrestricted done, QI failed). Re-upload only the failed quantity.
8. Not compiled or tested on a SAP system - needs syntax check (SE38) and the 15 test scenarios.
