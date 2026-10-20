# RelaxCat Derived Architecture Decisions

These are agent-revisable implementation decisions derived from confirmed product intent. They are not human intent by themselves.

## D-ARCH-001

Status: Active

Upstream: S-CAPTURE-001, S-CAPTURE-002, S-SIZE-001, S-EXCEL-001

Decision: Keep image sizing/layout logic independent from Avalonia, Greenshot, Win32 hooks, and Excel COM. Platform integrations are adapters around a testable core.

Reason: The confirmed behavior does not require UI or platform APIs to participate in pure sizing/layout calculations, and separating them reduces coupling and makes agent changes easier to verify.

Evidence: `docs/technical-design.md`, `src/RelaxCat.Core/`.

## D-RULE-001

Status: Active

Upstream: S-SIZE-002

Decision: Implement the scale language as a small deterministic parser/evaluator rather than executing arbitrary C# or another general scripting language.

Reason: The confirmed language is intentionally small and ordered. A dedicated evaluator is safer, easier to cache, and easier to test. Grammar details not explicitly confirmed may evolve without changing S-SIZE-002.

Evidence: `src/RelaxCat.Core/Imaging/ScaleRuleEngine.cs`.

## D-LAYOUT-001

Status: Active

Upstream: S-LAYOUT-001

Decision: Use deterministic shelf/row packing for space-saving layout rather than general bin packing.

Reason: It directly implements the confirmed left-to-right, then next-row behavior and keeps Excel output predictable.

Evidence: `src/RelaxCat.Core/Layout/ImageLayoutEngine.cs`.

## D-CACHE-001

Status: Active

Upstream: S-CACHE-001, S-SIZE-001, S-SIZE-002, S-SIZE-003

Decision: Cache scale results by effective image metadata plus effective sizing configuration/rule identity. Cache invalidation follows key changes rather than treating cached values as durable intent.

Reason: Any input that can alter scale must distinguish cache entries.

Evidence: `src/RelaxCat.Core/Imaging/ScaleCache.cs`.

## D-EXCEL-001

Status: Proposed

Upstream: S-EXCEL-001, S-EXCEL-002, S-EXCEL-003

Decision: Implement Excel 2016 live automation behind an Excel infrastructure boundary, keeping raw COM objects out of Avalonia ViewModels and core domain models.

Reason: The required live workbook reuse, worksheet normalization, shape insertion, and `xlMoveAndSize` behavior fit Office automation, while isolation contains COM lifecycle concerns.

Evidence: `docs/technical-design.md`. Implementation not yet present.
