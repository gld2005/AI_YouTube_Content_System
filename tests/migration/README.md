# Migration Test Fixture

`baseline_manifest.json` is the immutable expectation set for the baseline workbook.

Before a migration run:

1. Verify the source SHA-256.
2. Verify the source workbook contains the nine expected sheets.
3. Verify the seed-record counts meet or exceed the documented minimums.
4. Scan the baseline and output workbook for formula errors.

After a migration run:

1. Compare record counts and IDs against the crosswalk.
2. Verify the migration invariants in the manifest.
3. Record exceptions without overwriting source values.
4. Save a human-readable summary and test log in English.

The fixture intentionally uses minimum counts. It supports future user data additions without requiring edits to the baseline expectation set.
