# RelaxCat Singularity Entry

Read [config.md](config.md) for the pinned protocol and adoption boundary.

For each task:

1. Read only the confirmed intent records relevant to the requested area.
2. Inspect related implementation, tests, and documents as evidence; expand context only when needed.
3. Derive implementation details that do not require a product/intent choice.
4. Never infer a new confirmed intent from existing code or documentation.
5. Expose real conflicts and pause only the part that depends on the unresolved choice.
6. When confirmed intent changes, reevaluate downstream decisions and affected artifacts.

Project intent starts in [singularities](singularities/). Important derived architecture choices belong in [decisions](decisions/architecture.md).

Only `Confirmed` singularities are authoritative project intent. The technical design and current implementation are downstream evidence and may be corrected when they disagree with confirmed intent.

## CLI and local editor

With the pinned Singularity checkout alongside RelaxCat, run these commands from the RelaxCat root:

```powershell
python ../singularity/tools/singularity.py validate
python ../singularity/tools/singularity.py serve
```

Open the loopback address printed by the editor. Each TOML file contains one intent record. For records migrated from 0.1.0, `actor = "AI"` identifies the format migration; the original human confirmation and retirement bases are recorded in `reason`.

The star map reads optional Chinese display names from [node-names.toml](node-names.toml). Edit the value for a stable ID and click 「重新加载」 in the editor to see the change. This feature uses the updated local editor.
