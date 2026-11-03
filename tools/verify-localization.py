"""Validate external language packs, composite-format parameters and literal source keys."""
from pathlib import Path
import json
import re

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "src/RelaxCat.App"

def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"Duplicate resource key: {key}")
        result[key] = value
    return result

packs = {}
for path in sorted((APP / "Localization").glob("*.json")):
    pack = json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_object)
    culture = pack.get("culture")
    if culture != path.stem or not isinstance(pack.get("displayName"), str) or not pack["displayName"].strip():
        raise ValueError(f"{path.name}: culture must match filename; displayName is required")
    if not isinstance(pack.get("sortOrder", 100), int) or not isinstance(pack.get("strings"), dict):
        raise ValueError(f"{path.name}: invalid sortOrder or strings")
    if culture in packs:
        raise ValueError(f"Duplicate culture: {culture}")
    for key, value in pack["strings"].items():
        if not isinstance(value, str) or not value.strip():
            raise ValueError(f"{culture}: {key} must be a nonempty string")
    packs[culture] = pack["strings"]
baseline = packs["zh-CN"]
for culture, entries in packs.items():
    if entries.keys() != baseline.keys():
        raise ValueError(f"{culture}: missing {baseline.keys() - entries.keys()}, extra {entries.keys() - baseline.keys()}")
    for key, value in entries.items():
        parameters = lambda text: set(re.findall(r"(?<!\{)\{(\d+)(?:[,}:])", text))
        if parameters(value) != parameters(baseline[key]):
            raise ValueError(f"{culture}: inconsistent format parameters for {key}")
for path in APP.rglob("*"):
    if path.suffix not in (".cs", ".axaml") or {"bin", "obj"} & set(path.parts):
        continue
    source = path.read_text(encoding="utf-8-sig")
    keys = re.findall(r'(?:Text|text)\["([^"\n]+)"\]', source)
    if path.suffix == ".axaml":
        keys += re.findall(r'Text\[([A-Za-z][A-Za-z0-9_]*)\]', source)
    for key in keys:
        if key not in baseline:
            raise ValueError(f"{path.relative_to(ROOT)}: missing literal resource {key}")
print(f"PASS: {len(packs)} packs, {len(baseline)} keys per language, format parameters and literal references")
