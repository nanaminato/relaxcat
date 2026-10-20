# RelaxCat Singularity Entry

Read [config.md](config.md) for the pinned protocol and adoption boundary.

For each task:

1. Read only the confirmed intent records relevant to the requested area.
2. Inspect related implementation, tests, and documents as evidence; expand context only when needed.
3. Derive implementation details that do not require a product/intent choice.
4. Never infer a new confirmed intent from existing code or documentation.
5. Expose real conflicts and pause only the part that depends on the unresolved choice.
6. When confirmed intent changes, reevaluate downstream decisions and affected artifacts.

Project intent starts in [singularities](singularities/core.md). Important derived architecture choices belong in [decisions](decisions/architecture.md).

Only `Confirmed` singularities are authoritative project intent. The technical design and current implementation are downstream evidence and may be corrected when they disagree with confirmed intent.
