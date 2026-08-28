# Parked Issue: LCB Footer Bidder Name

## Status
Parked. Do not merge the `test/fix-lcb-footer-biddersname` branch into `main` unless the footer issue is resolved and tested.

## Working baseline
`main` is the production baseline for LCB document generation and should remain unchanged.

The current LCB workflow successfully:
- accepts the Notice template by Google Docs URL;
- accepts the destination folder by Google Drive URL;
- preserves BAC Secretariat project-specific template fields;
- writes `{{BiddersName}}` in the document body;
- writes `{{BidTable}}` into the notice;
- formats quantities and monetary values appropriately;
- generates bidder-specific notices; and
- provides a clickable link to the destination folder.

## Remaining issue
The template footer contains a 2-row × 2-column table. The bidder name placeholder is in **Row 1, Column 1** (zero-based row 0, column 0), in the same cell as the project title:

`(Project Title) | {{BiddersName}}`

The footer is configured to be the same throughout the document.

The current experimental implementation targets the footer table cell directly, but testing still does not replace `{{BiddersName}}`.

## Next investigation
When revisiting this issue, inspect the actual Google Docs footer/table text structure rather than changing the working LCB generation workflow. In particular, determine whether the placeholder is split across text runs or represented by another element inside the table cell, and identify a reliable way to replace it while preserving the existing footer formatting.

## Important
This issue is intentionally isolated from `main`. The footer is a minor remaining manual step and should not justify destabilizing the working LCB generator.