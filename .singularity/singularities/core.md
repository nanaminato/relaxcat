# RelaxCat Confirmed Singularities

These records capture product intent explicitly stated by the human during the RelaxCat design conversation. They constrain downstream implementation but intentionally do not prescribe ordinary code structure.

## S-CAPTURE-001

Status: Confirmed

Statement: RelaxCat supports selectable screenshot providers. Greenshot is a bundled executable provider; when selected, RelaxCat manages the Greenshot process when starting, switching providers, and exiting.

Confirmation basis: Human clarification, 2026-10-20.

## S-CAPTURE-002

Status: Confirmed

Statement: For configured monitored keys, RelaxCat may move the mouse to a configured position before the same key continues to Greenshot; RelaxCat's monitoring must not prevent Greenshot from processing that key.

Confirmation basis: Human clarification, 2026-10-20.

## S-IMPORT-001

Status: Confirmed

Statement: Screenshots produced through the monitored capture folder and images imported manually enter the RelaxCat image/export workflow. Long import and export operations expose progress and can be cancelled.

Confirmation basis: Human project description, 2026-10-20.

## S-SIZE-001

Status: Confirmed

Statement: RelaxCat does not normally resample image pixels for Excel layout. It changes the displayed image size, and DPI participates in sizing so images with different DPI metadata have more consistent displayed results.

Confirmation basis: Human clarification, 2026-10-20.

## S-SIZE-002

Status: Confirmed

Statement: Image scaling first evaluates user-defined ordered rules. Conditions are single-line expressions that may combine comparisons with `and` and `or`; rules are tried in order and the first match supplies the scale. If no rule matches, sizing falls back to the configured design size and design scale.

Confirmation basis: Human clarification, 2026-10-20.

## S-SIZE-003

Status: Confirmed

Statement: Fallback sizing preserves aspect ratio, approaches the configured design-size/design-scale target without exceeding its width or height, and limits enlargement so small images are not excessively enlarged.

Confirmation basis: Human clarification, 2026-10-20.

## S-CACHE-001

Status: Confirmed

Statement: Scale calculation should be cacheable because a given effective image size and sizing configuration normally produce a deterministic scale; the cache exists to avoid repeated calculation.

Confirmation basis: Human clarification, 2026-10-20.

## S-LAYOUT-001

Status: Confirmed

Statement: Default Excel layout places one image per row. In space-saving mode, a row still has a maximum usable width derived from design size and design scale; smaller images continue horizontally with configured spacing until the next image no longer fits, then layout continues on the next row.

Confirmation basis: Human clarification, 2026-10-20.

## S-EXCEL-001

Status: Confirmed

Statement: Excel export reuses an available document/session until it is no longer usable. If no usable document exists, RelaxCat creates a new document from the configured template.

Confirmation basis: Human project description and clarification, 2026-10-20.

## S-EXCEL-002

Status: Confirmed

Statement: Before exporting images, RelaxCat normalizes the target worksheet according to configuration, including row heights and column widths. This normalization occurs for both newly created and already available documents.

Confirmation basis: Human clarification, 2026-10-20.

## S-EXCEL-003

Status: Confirmed

Statement: Exported Excel images are configured to move and resize with cells (`xlMoveAndSize`).

Confirmation basis: Human project description, 2026-10-20.
