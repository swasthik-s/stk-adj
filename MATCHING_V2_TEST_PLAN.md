# Matching v2: test and rollout plan

## What changed

- Uses description tokens plus normalized size features, explicit pack quantity, and Brand/Class/Group as ranking evidence.
- Imports the master-list fields Pack Qty, Packing, and IsBundleItem when present.
- Rejects a candidate when both descriptions contain a size and those sizes conflict.
- Requires positive stock and an active item for parent/swap candidates.
- Uses explicit pack quantity for bundle conversion; it does not infer a conversion from generic words such as box or offer.
- Flags mixed-size/combo descriptions for manual review instead of suggesting a one-target bundle break.
- Adds a why_match explanation to suggested rows. Similarity is a candidate-ranking signal, not proof that two barcodes are interchangeable.

## Run tests

From the repository root:

\`\`\`bash
pip install pandas pytest streamlit openpyxl altair
pytest -q
\`\`\`

These focused tests extract the matching section without importing/running the Streamlit page.

## Before merging to main

1. Run the test suite and fix any failures.
2. Test with a copy of the real ItemSearchList.csv and NegativeStockCategorywise.xlsx.
3. Manually review a sample of bundle-break and wrong-sale suggestions, including wrong-size products, bundle products, inactive products, and low-stock parents.
4. Compare precision (correct suggestions / suggestions reviewed) and recall (known correct cases found / known correct cases) against the current version.
5. Do not automatically post adjustments from a fuzzy match. Keep the operator's confirmation step.

## Deployment

A branch commit does not update Streamlit if the app deployment is configured to track main. After validation, merge this branch into the configured deployment branch or explicitly change the Streamlit app's branch setting. Confirm the configured branch in Streamlit Cloud before assuming the live app has changed.
