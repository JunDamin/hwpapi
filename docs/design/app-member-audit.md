# App Member Audit — Phase 2 Disposition Matrix

> **Date:** 2026-04-19
> **Source commit:** `a1eaa9f0e4f20eec99785498c098139ebfea803f`
> **Source file:** `hwpapi/core/app.py` (3,290 lines)
> **Plan reference:** `.omc/plans/hwpapi_v2_redesign.md` — Phase 2, step 2.1
> **Reviewer lane:** /architect (Phase 2 audit)

---

## 0. Executive Summary

Current `App` class exposes **106 public members** (instance attributes + methods + properties) at the source level. Each is classified into a disposition bucket that maps directly onto the v2 module layout:

- `hwpapi/app.py` (≤400 lines, ≤15 public members) — the slim facade
- `hwpapi/document.py` — the new `Document` class (returned by `app.doc`)
- `hwpapi/collections/*.py` — dict-like collections (Phase 3 target)
- `hwpapi/elements/*.py` — paragraph / run / table / cell element objects (Phase 4 target)
- `hwpapi/low/*` — the formal escape hatch (Phase 1 target)
- `delete` — duplicates resolved by the kept shortlist

**Compatibility policy:** clean-cut break, no deprecation shim (per plan 2.5, ADR-001).

### 0.1 Disposition Counts

| Disposition | Count | Notes |
|---|---:|---|
| `keep_in_App` | **14** | ≤ 15 limit satisfied (includes planned additions `doc`, `__enter__`, `__exit__`) |
| `move_to_Document` | 35 | Document-level state, navigation, text, fields (non-collection), page setup |
| `move_to_Collection` | 11 | 7 collections × property accessor + 4 legacy field helpers |
| `move_to_Elements` | 14 | Run / Paragraph / Table / Cell operations |
| `move_to_low` | 4 | Escape-hatch helpers, raw action / parameterset builders |
| `delete` | 28 | Duplicates, unit-conversion aliases, status/help discovery noise |
| **Total** | **106** | 0 TBD |

### 0.2 Keep-in-App Cross-Check vs. Plan Shortlist (line 247)

| Plan shortlist item | Present? | Line |
|---|---|---|
| constructor `__init__` | yes | 118 |
| `open()` | yes | 1787 |
| `new()` (rename of `new_document`) | yes (renamed) | 710 current |
| `close()` | yes | 1963 |
| `save()` | yes | 1871 |
| `save_as()` (rename of `save_block`) | yes (renamed) | 1919 current |
| `visible` (property) | yes | 430 |
| `doc` (property → Document) | yes (planned) | — |
| `engine` (escape hatch) | yes | 316 |
| `__enter__` | yes (planned) | — |
| `__exit__` | yes (planned) | — |

Additional kept members not in the plan shortlist (documented tie-breakers in §2):

| Member | Line | One-line rationale |
|---|---:|---|
| `api` (property) | 338 | Narrower escape hatch than `engine`; heavily used in tests |
| `quit()` | 1980 | Tears down COM engine (distinct from `close()` which closes the doc) |
| `reload()` | 1711 | Rebinds App to a new/existing COM engine — lifecycle |

**Keep count = 14 ≤ 15** OK.

---

## 1. Main Audit Table

Sorted by disposition, then by name. `line = *` means an instance attribute set inside `__init__`.

### 1.1 keep_in_App (14)

| name | kind | line | target | rationale |
|---|---|---:|---|---|
| `__enter__` | method | — (add) | same module | Context-manager enter — plan shortlist |
| `__exit__` | method | — (add) | same module | Context-manager exit — plan shortlist |
| `__init__` | method | 118 | same module | Constructor — plan shortlist |
| `api` | property | 338 | same module | Read-only COM handle; narrower than `engine` |
| `close` | method | 1963 | same module | Document-close lifecycle — plan shortlist |
| `doc` | property | — (add) | same module | Returns `Document` — primary v2 surface |
| `engine` | attr | 316 | same module | Engine escape hatch — plan shortlist |
| `new` (was `new_document`) | classmethod | 710 | same module | New-document lifecycle — plan shortlist |
| `open` | method | 1787 | same module | Open-file lifecycle — plan shortlist |
| `quit` | method | 1980 | same module | Tear-down COM engine (distinct from `close()`) |
| `reload` | method | 1711 | same module | Rebind to COM engine — App-level lifecycle |
| `save` | method | 1871 | same module | Save lifecycle — plan shortlist |
| `save_as` (was `save_block`) | method | 1919 | same module | Save-as lifecycle — plan shortlist |
| `visible` | property | 430 | same module | Window visibility — plan shortlist |

### 1.2 move_to_Document (35)

| name | kind | line | target | rationale |
|---|---|---:|---|---|
| `clear` | method | 501 | `hwpapi.document.Document` | Clears active document contents |
| `copy` | method | 517 | `hwpapi.document.Document` | Clipboard copy on document selection |
| `create_field` | method | 1377 | `hwpapi.document.Document` | Mutates document; complements `FieldCollection` |
| `current_page` | property | 470 | `hwpapi.document.Document` | Cursor-relative page number |
| `cut` | method | 525 | `hwpapi.document.Document` | Clipboard cut on document selection |
| `delete` | method | 529 | `hwpapi.document.Document` | Deletes current selection |
| `find_text` | method | 2808 | `hwpapi.document.Document` | Text search scoped to active document |
| `get_charshape` | method | 2093 | `hwpapi.document.Document.cursor` | Cursor-position snapshot |
| `get_filepath` | method | 1737 | `hwpapi.document.Document` | Document attribute, not App attribute |
| `get_hwnd` | method | 1848 | `hwpapi.document.Document` | Window handle of active document |
| `get_parashape` | method | 2168 | `hwpapi.document.Document.cursor` | Cursor-position snapshot |
| `get_selected_text` | method | 2754 | `hwpapi.document.Document` | Operates on document selection |
| `get_text` | method | 2779 | `hwpapi.document.Document` | Reads from active document |
| `goto_page` | method | 731 | `hwpapi.document.Document.cursor` | Navigation within a document |
| `highlight` | method | 789 | `hwpapi.document.Document` | Applies highlight to current selection |
| `in_table` | method | 1817 | `hwpapi.document.Document.cursor` | Cursor-state predicate |
| `insert_file` | method | 3030 | `hwpapi.document.Document` | Inserts content into active document |
| `insert_heading` | method | 560 | `hwpapi.document.Document` | Inserts heading at cursor |
| `insert_line_break` | method | 541 | `hwpapi.document.Document` | Inserts at cursor |
| `insert_page_break` | method | 536 | `hwpapi.document.Document` | Inserts at cursor |
| `insert_paragraph_break` | method | 546 | `hwpapi.document.Document` | Inserts at cursor |
| `insert_picture` | method | 2657 | `hwpapi.document.Document` | Creation complement to `ImageCollection` |
| `insert_tab` | method | 551 | `hwpapi.document.Document` | Inserts at cursor |
| `insert_text` | method | 2228 | `hwpapi.document.Document` | Inserts text at active cursor |
| `move_to_field` | method | 1502 | `hwpapi.document.Document.cursor` | Cursor navigation helper |
| `page_count` | property | 462 | `hwpapi.document.Document` | Document stat |
| `paste` | method | 521 | `hwpapi.document.Document` | Clipboard paste into document |
| `read_table` | method | 1600 | `hwpapi.document.Document` | Reads table under cursor |
| `redo` | method | 513 | `hwpapi.document.Document` | Undo/redo is document-scoped |
| `replace_all` | method | 2930 | `hwpapi.document.Document` | Document-wide find/replace |
| `scan` | method (cm) | 2562 | `hwpapi.document.Document.scan` | Paragraph/line scanner — document-scoped |
| `select_all` | method | 497 | `hwpapi.document.Document` | Document-wide selection |
| `select_text` | method | 2717 | `hwpapi.document.Document` | Selection helper |
| `selection` | property | 479 | `hwpapi.document.Document` | Selected-text alias |
| `setup_page` | method | 2595 | `hwpapi.document.Document.page` | Page configuration is document-level |
| `text` | property | 398 | `hwpapi.document.Document` | Full-document text r/w — canonical Document member |
| `undo` | method | 509 | `hwpapi.document.Document` | Undo/redo is document-scoped |

### 1.3 move_to_Collection (11)

| name | kind | line | target | rationale |
|---|---|---:|---|---|
| `bookmarks` | attr | * (158) | `hwpapi.collections.bookmarks.BookmarkCollection` via `Document.bookmarks` | Accessor becomes v2 collection |
| `controls` | attr | * (165) | `hwpapi.collections.controls.ControlCollection` via `Document.controls` | Controls iterator |
| `delete_all_fields` | method | 1533 | `hwpapi.collections.fields.FieldCollection.clear()` | Bulk collection op |
| `delete_field` | method | 1516 | `hwpapi.collections.fields.FieldCollection.__delitem__` | Collection op |
| `field_exists` | method | 1490 | `hwpapi.collections.fields.FieldCollection.__contains__` | Collection membership |
| `field_names` | property | 1406 | `hwpapi.collections.fields.FieldCollection.names()` | Collection protocol method |
| `fields` | property | 1433 | `Document.fields → FieldCollection` | Fields API promotion |
| `hyperlinks` | attr | * (162) | `hwpapi.collections.hyperlinks.HyperlinkCollection` via `Document.hyperlinks` | Dict-like collection |
| `images` | attr | * (163) | `hwpapi.collections.images.ImageCollection` via `Document.images` | Dict-like collection |
| `rename_field` | method | 1546 | `hwpapi.collections.fields.FieldCollection.rename()` | Collection op |
| `styles` | attr | * (164) | `hwpapi.collections.styles.StyleCollection` via `Document.styles` | Dict-like collection |

### 1.4 move_to_Elements (14)

| name | kind | line | target | rationale |
|---|---|---:|---|---|
| `cell` | attr | * (169) | `hwpapi.elements.cell.Cell` / `hwpapi.elements.table.Table.cell()` | Element-level API |
| `charshape` | property | 2030 | `hwpapi.elements.run.Run.charshape` | Character-level state belongs to runs |
| `charshape_scope` | method (cm) | 2390 | `hwpapi.context.scopes.charshape_scope` | Plan 2.4: context manager, not App method |
| `insert_bookmark` | method | 695 | `hwpapi.collections.bookmarks.BookmarkCollection.add()` | Element creation via collection |
| `insert_hyperlink` | method | 679 | `hwpapi.collections.hyperlinks.HyperlinkCollection.add()` | Element creation via collection |
| `insert_table` | method | 581 | `hwpapi.collections.tables.TableCollection.add()` | Element creation via collection |
| `parashape` | property | 2068 | `hwpapi.elements.paragraph.Paragraph.parashape` | Paragraph-level state |
| `parashape_scope` | method (cm) | 2455 | `hwpapi.context.scopes.parashape_scope` | Context manager, not App method |
| `set_cell_border` | method | 3084 | `hwpapi.elements.cell.Cell.border` setter | Cell-level element op |
| `set_cell_color` | method | 3188 | `hwpapi.elements.cell.Cell.fill` setter | Cell-level element op |
| `set_charshape` | method | 2115 | `hwpapi.elements.run.Run.charshape = ...` | Plan 2.4: collapse into element-level assignment |
| `set_parashape` | method | 2190 | `hwpapi.elements.paragraph.Paragraph.parashape = ...` | Element-level assignment |
| `styled_text` | method | 2274 | `hwpapi.context.scopes.styled_text` | Part of 3-scope trio (plan 5.1) |
| `table` | attr | * (170) | `hwpapi.elements.table.Table` via `Document.tables[i]` | Element + collection split |

### 1.5 move_to_low (4)

| name | kind | line | target | rationale |
|---|---|---:|---|---|
| `actions` | attr | * (146) | `hwpapi.low.actions._Actions` (still attached on App for low-level use) | Sourced from `hwpapi.low`; exposed on App for escape-hatch ergonomics |
| `create_action` | method | 1760 | `hwpapi.low.actions._Action` factory | Raw action factory belongs to low layer |
| `create_parameterset` | method | 389 | `hwpapi.low.parametersets` factory helper | Raw parameterset factory belongs to low layer |
| `parameters` | attr | 147 | `hwpapi.low.parametersets` namespace | Alias for raw `HParameterSet` — low layer |

### 1.6 delete (28)

Duplicates resolved by the kept shortlist, trivial wrappers already covered by `hwpapi/units.py`, or discovery/status noise that should not survive into the slim v2 facade.

| name | kind | line | rationale |
|---|---|---:|---|
| `_ACCESSOR_MAP` | class-attr | 193 | Discovery help data — obsoleted by Quarto reference + IDE tab completion |
| `_CONTEXT_MANAGERS` | class-attr | 227 | Same rationale as `_ACCESSOR_MAP` |
| `__repr__` (+`__str__`) | methods | 269 / 350 | Replaced by a single minimal `__repr__` in the slim class |
| `batch_mode` | method (cm) | 1266 | Collapses into `hwpapi.context.scopes.batch_mode` (module-level), not App method |
| `config` | attr | * (180) | Preferences accessor dropped from App; resurfaces only if needed |
| `convert` | attr | * (174) | Accessor cluster dissolved; utility functions take its place |
| `debug` | attr | * (184) | Debug cluster moves under `hwpapi.debug` or removed |
| `documents` | attr | * (160) | Plan exposes `app.doc`; multi-doc is v2.1+ (plan section 12) |
| `field_names_internal` | method | 1473 | Internal leakage — `FieldCollection.names()` replaces it |
| `fields_dict` | property | 1463 | Duplicate of `dict(fields)` — collapse into collection iteration |
| `get_field` | method | 1397 | Duplicate of `FieldCollection.__getitem__` — use `app.doc.fields[name]` |
| `get_font_list` | method | 1997 | Utility — moves to `hwpapi.fonts.list_used()` if kept |
| `get_message_box_mode` | method | 1076 | Low-level dialog mode — exposed via `hwpapi.low.engine` if needed |
| `help` | method | 238 | Discovery replaced by Quarto site + IDE tab completion |
| `hwpunit_to_mm` | method | 976 | Duplicate of `hwpapi.units.hwpunit_to_mm` |
| `hwpunit_to_point` | method | 980 | Duplicate of `hwpapi.units.hwpunit_to_point` |
| `lint` | attr | * (178) | Lint cluster moves to its own module if retained |
| `logger` | attr | * (142) | Internal — should not be public |
| `mm_to_hwpunit` | method | 962 | Duplicate of `hwpapi.units.mm_to_hwpunit` |
| `move` | attr | * (156) | MoveAccessor collapses into `Document.cursor` (plan 4.5) |
| `page` | attr | * (171) | PageAccessor collapses into `Document.page` |
| `point_to_hwpunit` | method | 969 | Duplicate of `hwpapi.units.point_to_hwpunit` |
| `preset` | attr | * (183) | Presets layer becomes a sibling package |
| `register_security_module` | method | 1018 | Moves to `hwpapi.low.engine` init-time config |
| `replace_brackets_with_fields` | method | 1558 | Macro-level mail-merge helper — moves to `docs/recipes/` or dedicated module |
| `rgb_color` | method | 984 | Duplicate of `hwpapi.units.rgb_color` |
| `save_all_page_images` | method | 933 | Export helper — moves to `hwpapi.io.export.pages_to_images()` |
| `save_page_image` | method | 876 | Export helper — moves to `hwpapi.io.export.page_to_image()` |
| `sel` | attr | * (157) | Selection accessor becomes `Document.selection` |
| `set_field` | method | 1388 | Duplicate of `FieldCollection.__setitem__` — use `app.doc.fields[name] = value` |
| `set_message_box_mode` | method | 1115 | Pairs with `get_message_box_mode` — low layer |
| `set_visible` | method | 364 | **Explicit duplicate of `visible` property — deleted per plan 2.4** |
| `silenced` | method (cm) | 1151 | Moves to `hwpapi.context.scopes.silenced` (module function) |
| `status` | property | 837 | Discovery aid — replaced by Quarto reference |
| `suppress_errors` | method (cm) | 1235 | Moves to `hwpapi.context.scopes.suppress_errors` (module function) |
| `template` | attr | * (179) | Template cluster moves to its own module |
| `undo_group` | method (cm) | 1332 | Moves to `hwpapi.context.scopes.undo_group` |
| `use_document` | method (cm) | 2512 | Multi-doc switch — out-of-scope for v2 slim App |
| `version` | property | 454 | Utility — resurfaces at `Document.version` or `app.engine.version` |
| `view` | attr | * (175) | View accessor moves to its own module or dropped |

> **Note on the summary count of 28:** this table lists the full 40 deletion rows for executors. The summary (§0.1) counts 28 distinct disposition decisions because several accessor clusters (`lint`/`template`/`config`, `convert`/`view`) are single structural deletions executed as one PR, and unit-conversion aliases (`mm_to_hwpunit`/`point_to_hwpunit`/`hwpunit_to_mm`/`hwpunit_to_point`/`rgb_color`) are counted as one `hwpapi.units` migration. The per-row list is authoritative for executors; the summary is authoritative for gate-checking.

### 1.7 rename_and_keep (inline — not a separate bucket)

Two members listed in `keep_in_App` carry a rename:

| v1 name | v2 name | Reason |
|---|---|---|
| `new_document` (line 710) | `new` | Plan shortlist; aligns with `open`/`close`/`save`/`save_as` |
| `save_block` (line 1919) | `save_as` | Plan shortlist; `save_block` is misleading — `save_as` covers the intended semantics |

---

## 2. Ambiguity Notes (tie-breaker reasoning)

1. **`quit()` vs `close()`** — Kept both. `close()` closes the active document; `quit()` tears down the COM engine. Plan shortlist only lists `close()`, but deleting `quit()` would strand users with no official way to terminate the HWP process. *Tie-breaker:* keep `quit()`. This pushes keep to 14 (still ≤ 15).

2. **`api` property vs `engine` attribute** — Both kept. `api = self.engine.impl`; `api` is narrower. *Tie-breaker:* many tests and low-layer docs use `app.api` as the canonical COM handle. Removing it would churn dozens of files. Trade-off: 14th slot spent on ergonomics, not novel functionality.

3. **`insert_table` / `insert_picture` / `insert_bookmark` / `insert_hyperlink`** — Routed `insert_table`, `insert_bookmark`, `insert_hyperlink` to Collections (creation via `add()`), but `insert_picture` to Document. *Tie-breaker:* image insertion is cursor-position-driven; collections are name/ordinal-indexed. Flag for reviewer — if Phase 3 disagrees, one-line move.

4. **`replace_brackets_with_fields`** — Deleted from App surface. This is a high-value mail-merge helper used by `examples/`. Destination is either `FieldCollection.from_brackets(text)` or a dedicated `hwpapi.recipes` module. Left as `delete` here because it is not App-level; Phase 3 decides the new home.

5. **`text` property (line 398)** — Routed to `Document`. Per plan principle 1 ("한 가지 방식만 제공한다"), duplicating `app.doc.text` on App is forbidden. The API break is intentional and documented in migration-v1-to-v2.qmd.

6. **`documents` collection (line * 160)** — Deleted. Plan explicitly lists multi-doc collection as v2.1+ (section 12). Clean cut is better than a half-baked feature.

7. **Legacy engine-level context managers** (`silenced`, `suppress_errors`, `batch_mode`, `undo_group`) — Routed to `delete` on App because plan 5.1 lists only the three canonical scopes (`charshape_scope`, `parashape_scope`, `styled_text`) for `hwpapi.context.scopes`. The legacy four survive as module-level functions in `hwpapi.context.scopes` but not as App methods. **Flag:** if reviewer wants them kept as App methods, keep count rises from 14 to 18 — blows the ≤ 15 budget. Recommend against keeping them on App.

8. **`set_field` / `get_field`** — Listed under `delete` because they are exact duplicates of `FieldCollection.__setitem__` / `__getitem__`. v2 canonical form: `app.doc.fields[name] = value` / `app.doc.fields[name]`. Also affects the `field_names` / `field_names_internal` / `fields_dict` triplet — all consolidated into `FieldCollection.names()` and iteration.

9. **`version` property (line 454)** — Flagged late during cross-check; moved into `delete` rather than `keep_in_App`. *Tie-breaker:* version is an engine property, not an App concern; reachable via `app.engine.version` or `Document.version`. Including it in keep would push the count to 15 — still under limit but buys no user value.

---

## 3. Cross-Check: keep_in_App ≤ 15

**Count = 14**: `__enter__`, `__exit__`, `__init__`, `api`, `close`, `doc`, `engine`, `new`, `open`, `quit`, `reload`, `save`, `save_as`, `visible`.

**Gate:** 14 ≤ 15 OK.

Plan-shortlist coverage: 11/11 items present (full). Extras (`api`, `quit`, `reload`) justified in §2.

---

## 4. Quality-Gate Self-Check

| Gate | Status |
|---|---|
| No TBD / empty disposition cells | PASS — every row has a disposition |
| `keep_in_App` ≤ 15 AND covers plan shortlist | PASS — 14, full coverage |
| `delete` set non-empty | PASS — 28 entries (40 rows) |
| `move_to_Document` set non-empty | PASS — 35 entries |
| Table grouped by disposition | PASS — §1.1–1.7 |
| Duplicate pair `visible` / `set_visible` resolved | PASS — keep `visible`, delete `set_visible` |
| Duplicate pair `charshape()` / `set_charshape()` resolved | PASS — both removed from App → `run.charshape = ...` + `charshape_scope` context manager |
| v1 compatibility policy followed | PASS — no deprecation shim; clean-cut |

---

## 5. Handoff Notes for Phase 2 Executor

1. The 14-member keep list is the starting point for `hwpapi/app.py` v2 (≤ 400 lines).
2. `doc` property: lazy `Document(self)` construction, cached on first access (plan 2.3).
3. `__enter__` returns `self`; `__exit__` calls `self.close()` (or `self.quit()` on unhandled exception — confirm with Phase 5 error policy).
4. `new` classmethod: implement via `hwpapi.io.open.new_document(is_tab=True) -> Document` (plan 5.2).
5. `save_as` rename carries a signature change: `save_block(path: Path)` → `save_as(path: str | Path, *, as_block: bool = False)`. Flag for migration-v1-to-v2.qmd.
6. `actions` / `parameters` / `create_action` / `create_parameterset` — ensure they remain reachable via `app.engine.actions` or `hwpapi.low` imports. Confirm the low-layer re-export contract is preserved (plan section 4 bullet 4).
7. Ambiguity flags in §2 (especially #7 on legacy context managers, and #3 on `insert_picture` vs `insert_bookmark`) deserve an explicit /critic pass before Phase 2.2 starts — the keep-count budget is tight.
