# Matching v3: validation and rollout

The conservative matcher is now integrated into main through views/negative_stock.py and matching_engine.py.

## Local validation against the supplied reports

- Master list: 102,652 item rows.
- Negative-stock report: 1,565 negative rows.
- Focused regression tests: 8 passed locally.
- Candidate smoke test: 68 bundle-break suggestions, 139 manual-review rows, and 18 wrong-sale suggestions.

These are candidate counts, not accuracy scores. The reports do not contain verified labels for every negative line, so precision and recall cannot be calculated without an operator-reviewed ground-truth sample.

## Safety rules

- A shared word alone cannot create a match.
- Candidate class/department must agree; conflicting size, apparel size, model code, brand, or known variant terms are rejected.
- Explicit pack quantity is preferred. A metadata-only Pack Qty requires an exact normalized description and corroborating unit-cost ratio.
- A description pack quantity that conflicts with master Pack Qty is rejected.
- Weighed KG negatives, sets, mixed-size combos, and negatives that are already multipacks are sent to manual review or skipped rather than converted blindly.
- Wrong-sale matching uses stricter name overlap, stock, class, size, brand, variant, and selling-price checks.
- Suggestions never post adjustments automatically; staff must verify the barcode and conversion.

## Run regression tests

From the repository root, run pytest -q. GitHub Actions is configured to run the same test suite on relevant pushes and pull requests.

## Deployment note

The matching update is on main. Streamlit Community Cloud will use it only if the app's configured deployment branch is main; check the app's deployment settings and logs to confirm the live deployment succeeded.
