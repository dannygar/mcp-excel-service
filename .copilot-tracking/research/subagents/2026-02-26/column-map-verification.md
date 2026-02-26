# COLUMN_MAP Verification — `server.py` vs `core_operations.py`

**Date:** 2026-02-26
**Purpose:** Extract exact COLUMN_MAP definitions and cell-write logic from both files.

---

## 1. COLUMN_MAP in `core_operations.py` (lines 637–649)

```python
COLUMN_MAP = {
    "C": "open_date",
    "E": "open_time",
    "I": "strategy",
    "J": "credit",
    "K": "debit",
    "L": "contracts",
    "N": "open_fees",
    "O": "close_fees",
    "Q": "sold_call_strike",
    "R": "sold_put_strike",
    "T": "width",
}
```

## 2. COLUMN_MAP in `server.py` (lines 441–452)

```python
COLUMN_MAP = {
    "C": "open_date",
    "E": "open_time",
    "I": "strategy",
    "J": "credit",
    "K": "debit",
    "L": "contracts",
    "N": "open_fees",
    "O": "close_fees",
    "Q": "sold_call_strike",
    "R": "sold_put_strike",
    "T": "width",
}
```

### Verdict: IDENTICAL

Both files define the exact same 11 column mappings. No differences.

---

## 3. Cell-write loop in `core_operations.py` (lines 893–916)

```python
async with httpx.AsyncClient() as write_client:
    for col_letter, field_name in COLUMN_MAP.items():
        value = values_dict.get(field_name, "")
        if value == "" or value is None:
            continue

        cell_address = f"{col_letter}{target_row}"
        cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{cell_address}')"

        try:
            response = await write_client.patch(
                cell_url,
                headers=headers,
                json={"values": [[value]]},
                timeout=30.0
            )

            if response.status_code in [200, 201]:
                updated_cells.append(cell_address)
            else:
                error_data = response.json() if response.content else {}
                error_message = error_data.get("error", {}).get("message", response.text)
                cell_errors.append(f"{cell_address}: {error_message}")
        except Exception as cell_error:
            cell_errors.append(f"{col_letter}{target_row}: {str(cell_error)}")
```

## 4. Cell-write loop in `server.py` (lines 730–758)

```python
async with httpx.AsyncClient() as write_client:
    for col_letter, field_name in COLUMN_MAP.items():
        value = values_dict.get(field_name, "")
        # Skip empty values
        if value == "" or value is None:
            continue

        cell_address = f"{col_letter}{target_row}"
        cell_url = f"{workbook_url}/worksheets/{sheet_name}/range(address='{cell_address}')"

        try:
            response = await write_client.patch(
                cell_url,
                headers=headers,
                json={"values": [[value]]},
                timeout=30.0
            )

            if response.status_code in [200, 201]:
                updated_cells.append(cell_address)
            else:
                error_data = response.json() if response.content else {}
                error_message = error_data.get("error", {}).get("message", response.text)
                cell_errors.append(f"{cell_address}: {error_message}")
        except Exception as cell_error:
            cell_errors.append(f"{col_letter}{target_row}: {str(cell_error)}")
```

### Verdict: IDENTICAL

The cell-write loops are functionally identical (one has a `# Skip empty values` comment, otherwise the same).

---

## 5. `values_dict` construction (both files)

Both files build `values_dict` identically:

```python
values_dict = {
    "open_date": open_date,
    "open_time": open_time,
    "strategy": mapped_strategy,
    "credit": credit,
    "debit": debit,
    "contracts": trade.get("contracts", ""),
    "open_fees": open_fees,
    "close_fees": trade.get("close_fees", ""),
    "sold_call_strike": sold_call_strike if sold_call_strike else "",
    "sold_put_strike": sold_put_strike if sold_put_strike else "",
    "width": width if width else "",
}
```

---

## 6. Writes OUTSIDE the COLUMN_MAP loop

### In `logTrades` (both files): **NONE**

The **only** writes in `logTrades` happen inside the `for col_letter, field_name in COLUMN_MAP.items()` loop. No columns A, B, D, F, G, H, M, P, S, or any other column are written by `logTrades`.

### In other tools (for reference):

- **`closeTrade`** (`core_operations.py` lines 476–520): Writes to columns **F** (close_date) and **G** (close_time) directly, outside any COLUMN_MAP loop.
- **`updateTradeWithDelta`** (`core_operations.py` lines 290–310): Writes to a dynamic column determined by `DELTA_COLUMN_MAPPING` (based on time_window + strike_type).
- **`updateRange`** (`core_operations.py` / `server.py`): Writes to an arbitrary address passed as parameter.

---

## 7. Columns NOT written by `logTrades`

Based on the COLUMN_MAP, the following Excel columns are **never written** by `logTrades`:

| Column | Likely purpose | Written by |
|--------|---------------|------------|
| A | (unknown) | Nothing |
| B | (unknown) | Nothing |
| D | (unknown/gap) | Nothing |
| F | close_date | `closeTrade` |
| G | close_time | `closeTrade` |
| H | (unknown) | Nothing |
| M | (unknown/gap) | Nothing |
| P | (unknown/gap) | Nothing |
| S | (unknown/gap) | Nothing |

---

## Summary

- Both files have **identical** COLUMN_MAP definitions (11 entries: C, E, I, J, K, L, N, O, Q, R, T).
- Both files have **identical** cell-write loops (iterate COLUMN_MAP, PATCH each cell individually).
- Both files have **identical** `values_dict` construction.
- **No writes occur outside the COLUMN_MAP loop** in `logTrades`.
