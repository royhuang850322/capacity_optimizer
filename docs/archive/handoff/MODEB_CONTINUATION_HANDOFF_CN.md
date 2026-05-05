# ModeB Continuation Handoff

## 1. Purpose

This document is the handoff package for the next engineer/agent who will continue the ModeB redesign.

Scope of this handoff:

- Record what has already been completed in this round
- Freeze the business rules already confirmed by the owner
- Explain exactly what still needs to be implemented for the new ModeB
- Reduce the chance of reintroducing removed logic such as `Direct_Mode`


## 2. Completed In This Round

The following items are already implemented and verified:

1. `Direct_Mode` has been removed from runtime semantics
   - No more PQ/template input branch
   - Runtime now always reads from folder-based planner/master inputs
   - Launcher and control workbook no longer expose `Direct_Mode`

2. `ModeA` has been simplified to a single capacity basis
   - ModeA now uses only `master_capacity`
   - ModeA no longer derives `Max` / `Planner` from routing
   - ModeA output is now a single-basis workbook again

3. Documentation and visible UX were updated to match the above
   - Launcher
   - Control workbook template
   - User-facing docs that mentioned `Direct_Mode`

4. Targeted regression coverage is passing
   - Command used:
     - `python -m pytest -p no:cacheprovider tests/test_smoke_m8.py tests/test_desktop_launcher.py tests/test_regressions.py`
   - Result at handoff time:
     - `27 passed`


## 3. Current Code Baseline

Important files already changed in this round:

- `app/main.py`
- `app/data_loader.py`
- `app/models.py`
- `app/desktop_launcher.py`
- `app/create_template.py`
- `app/output_writer.py`
- `app/i18n.py`
- `tests/test_regressions.py`
- `tests/test_smoke_m8.py`
- `tests/test_desktop_launcher.py`

Current high-level behavior:

- ModeA:
  - fixed to folder-based input
  - single capacity basis
  - capacity-only solver

- ModeB:
  - still uses the old solver behavior for now
  - still runs dual basis output (`Max` / `Planner`)
  - still contains old family-aware routing behavior in the solver and validator
  - this is the next area to redesign


## 4. Rules Confirmed With Owner

These points are already confirmed and should be treated as frozen unless the owner explicitly changes them later.

### 4.1 Global Product / Family Rule

- `1 product -> 1 family`
- This rule should apply to both ModeA and ModeB
- `ProductFamily` is not a calculation key
- `ProductFamily` is reference/display metadata only
- In ModeB, `ProductFamily` must not be used for routing inheritance or solver routing eligibility
- If the same `product` appears with multiple `family` values in one run, it should be treated as data error

Implication:

- Current family-based routing behavior in the old ModeB implementation must be removed
- Current tolerant metadata merge behavior for conflicting family values should be reconsidered and replaced by validation failure

### 4.2 Toller Rule

- A product will not have multiple valid `Toller` options
- At most one eligible toller per product

### 4.3 Routing-Only Resource Rule

- A routing resource is allowed to exist only in routing
- It does not have to exist in `master_capacity`
- This is especially important for the Stage 2 overflow reallocation logic

### 4.4 Report Load Percentage Rule

- Do not force Stage 1 and Stage 2 allocations into one mixed load percentage denominator
- Keep allocation source visible and separate
- Prefer source-aware reporting over a misleading unified percentage

### 4.5 Intermediate Residual Rule

- The "intermediate residual" means the residual after Stage 1 capacity-only allocation
- Example:
  - after capacity stage remains `40`
  - after routing stage remains `15`
  - final classification turns `15` into `Outsourced` or `Unmet`
- Main report should show final result
- Intermediate residual should be preserved in trace/audit output, not confused with final result

### 4.6 Both Summary Baseline Rule

- Do not redesign the `Both` mode comparison baseline yet
- Keep that topic deferred until ModeB redesign is finished


## 5. New ModeB Target Logic

The owner wants ModeB to be redesigned from the current routing-first logic into a staged model.

### 5.1 Stage 1: Capacity Baseline

Run a first pass exactly like ModeA:

- ignore routing completely
- use capacity-based monthly ton allocation only
- use `master_capacity` as the baseline source
- maximize demand fulfillment first

Output of Stage 1:

- internal tons allocated by capacity baseline
- `Residual_After_Capacity`

### 5.2 Stage 2: Routing Overflow Reallocation

Only the residual from Stage 1 should enter this stage.

Routing rules for this stage:

- only product-level routing counts
- family-level routing is reference only and must not drive allocation
- if a product has no valid internal routing resource, it cannot be reallocated internally in Stage 2

Capacity rule for this stage:

- reallocation can only use capacity that remains after Stage 1
- if a routing resource has already been filled to 100%, it cannot take more

Optimization rule for this stage:

- still use solver optimization, not greedy row-by-row assignment
- primary objective:
  - maximize routed residual allocation / minimize remaining residual
- secondary objective:
  - prefer better priority

### 5.3 Stage 3: Final Classification

After Stage 2:

- if a product still has residual and has eligible toller:
  - classify that residual as `Outsourced`
- if no toller:
  - classify as `Unmet`


## 6. Max / Planned Comparison Rule For New ModeB

Working interpretation used in this handoff:

- Stage 1 uses a single baseline from `master_capacity`
- `Max` vs `Planned` comparison applies to Stage 2 routing overflow reallocation

Meaning:

- the first capacity-only baseline is shared
- the overflow routing stage is compared under:
  - `Max Capacity Ton`
  - `Planned Capacity Ton`

If the owner later says the comparison should start from Stage 1 as well, this is the first assumption to revisit.


## 7. Recommended Output Model For New ModeB

Recommended result structure:

- `Allocation_Source = Capacity_Base`
- `Allocation_Source = Routing_Reroute`
- `Allocation_Source = Toller`
- `Allocation_Source = Unmet`

Recommended process fields for audit/trace:

- `Residual_After_Capacity`
- `Residual_After_Routing`
- final classified result

Main report should prioritize:

- final internal tons
- final outsourced tons
- final unmet tons
- service level

Trace/audit sheets should preserve the intermediate stage values.


## 8. Files Most Likely To Change Next

### Core Logic

- `app/optimizer.py`
  - replace old ModeB 4-phase routing-first logic
  - implement Stage 1 + Stage 2 + Stage 3 flow

### Validation

- `app/validator.py`
  - remove family-based routing acceptance
  - enforce product-level routing expectations for new ModeB
  - add global `1 product -> 1 family` validation

### Data Loading / Semantics

- `app/data_loader.py`
  - parsing may stay similar
  - but downstream semantics for family use should no longer imply routing inheritance

### Reporting / Attribution

- `app/load_pressure.py`
  - current ModeB unmet/outsource attribution is based on old routing assumptions
  - must be updated to reflect the staged model

- `app/output_writer.py`
  - likely needs source-aware detail output
  - likely needs trace/audit staging sheets or staging columns

### Tests

- `tests/test_regressions.py`
- additional ModeB-specific tests should be added for the new staged behavior


## 9. Recommended Implementation Order

1. Add / tighten validation first
   - enforce `1 product -> 1 family`
   - make ModeB routing product-level only for solver eligibility

2. Refactor ModeB solver flow in `app/optimizer.py`
   - Stage 1 capacity baseline
   - Stage 2 routing overflow reallocation
   - Stage 3 toller/unmet classification

3. Preserve dual-basis compare behavior only around the overflow routing stage

4. Update report attribution logic
   - do not let old family-based or old primary/alternative assumptions survive in `load_pressure.py`

5. Update output model
   - allocation source
   - trace/audit support

6. Expand regression coverage


## 10. Minimum Test Matrix For The Next Round

At minimum, the next implementer should cover:

1. Capacity-only product
   - Stage 1 fully satisfies
   - Stage 2 should do nothing

2. Overflow product with valid product-level routing
   - Stage 1 leaves residual
   - Stage 2 reallocates using remaining routing resource capacity

3. Overflow product with no internal routing but with toller
   - residual should end as `Outsourced`

4. Overflow product with no internal routing and no toller
   - residual should end as `Unmet`

5. Product with routing-only resource not present in `master_capacity`
   - should still be eligible in Stage 2

6. Product with multiple family values in input
   - should fail validation

7. `Max` vs `Planned` overflow compare
   - shared Stage 1
   - different Stage 2 outcomes

8. Both mode should still run without redesigning the comparison baseline yet


## 11. Important Warning About The Current Repository State

The worktree is not fully clean.

There are unrelated modified/untracked items in the repo that were not part of this ModeA / handoff task.

Do not blindly revert unrelated changes.

In particular:

- review `git status` before making broad edits
- limit changes to files required for the ModeB redesign
- avoid resetting or cleaning unrelated user work


## 12. Suggested Prompt For The Next Engineer / Agent

Use this as the starting brief:

```text
Please continue the ModeB redesign in C:\Users\super\capacity_optimizer based on docs/MODEB_CONTINUATION_HANDOFF_CN.md.

Important constraints already confirmed by the owner:

1. Direct_Mode has already been removed. Do not reintroduce it.
2. ModeA is already finished and should stay capacity-only, single-basis.
3. ModeB must be redesigned into:
   - Stage 1: capacity-only baseline, same logic as ModeA
   - Stage 2: routing overflow reallocation only for Stage 1 residual
   - Stage 3: toller or unmet final classification
4. ProductFamily does not participate in routing inheritance or solver decisions.
5. Enforce 1 product -> 1 family across the run.
6. Routing reallocation must use solver optimization, not greedy allocation.
7. Toller is unique per product.
8. Routing-only resources are allowed.
9. Do not redesign Both summary baseline yet.

Before editing, inspect the current code and preserve the completed ModeA refactor.
After edits, run:
python -m pytest -p no:cacheprovider tests/test_smoke_m8.py tests/test_desktop_launcher.py tests/test_regressions.py
and extend tests for the new ModeB behavior.
```

