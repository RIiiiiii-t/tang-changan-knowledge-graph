from __future__ import annotations

import json
import re
import uuid
from pathlib import Path

import openpyxl

ROOT = Path.cwd()
OUT = ROOT / "outputs" / "019fdae0-9371-7c91-8389-5910cbea4548" / "volume1-2-platform-import"
NS = uuid.NAMESPACE_URL


def uid(kind: str, key: str) -> str:
    return str(uuid.uuid5(NS, f"tang-changan:{kind}:{key}"))


def records(book, index: int):
    sheet = book.worksheets[index]
    headers = [str(c.value or "").rstrip("*").strip() for c in sheet[3]]
    result = []
    for row in sheet.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        result.append({h: row[i] for i, h in enumerate(headers) if h})
    return result


def split(value):
    return [x.strip() for x in re.split(r"[,;；，\n]+", str(value or "")) if x.strip()]


def clean(obj):
    return {k: v for k, v in obj.items() if v not in (None, "", [])}


def status(value):
    return "published" if str(value or "").strip() in {"\u5ba1\u6838\u901a\u8fc7", "\u5df2\u53d1\u5e03"} else "draft"


v1 = next((ROOT / "07-\u534f\u4f5c\u5305").rglob("BATCH-0200_A_*\u5b8c\u6574\u5ba1\u6838\u7248*.xlsx"))
v2 = next((ROOT / "outputs" / "019fdae0-9371-7c91-8389-5910cbea4548").glob("BATCH-0201_A_*.xlsx"))
books = [openpyxl.load_workbook(p, read_only=True, data_only=True) for p in (v1, v2)]

source_rows = [r for b in books for r in records(b, 1)]
chunk_rows = [r for b in books for r in records(b, 2)]
entity_rows = [r for b in books for r in records(b, 3)]
relation_rows = [r for b in books for r in records(b, 4)]

sources = []
for r in source_rows:
    key = r["source_key"]
    sources.append({
        "id": uid("source", key), "title": r["title"], "source_type": r["source_type"],
        "identifier": key, "payload": clean({"project_key": key, "author": r.get("author"),
        "edition": r.get("edition"), "publisher": r.get("publisher"),
        "publication_year": r.get("publication_year"), "volume": r.get("volume"),
        "page": r.get("page"), "url_or_file": r.get("url_or_file"), "notes": r.get("notes")}),
        "status": status(r.get("review_status")), "visibility": "public", "confidence": 0.9,
        "data_classification": "tang_changan_curated"
    })

chunks = []
for r in chunk_rows:
    key = r["chunk_key"]
    chunks.append({
        "id": uid("chunk", key), "key": key, "source_id": uid("source", r["source_key"]),
        "original_text": r["content"], "payload": clean({"chapter": r.get("chapter"),
        "volume": r.get("volume"), "page": r.get("page"),
        "related_entity_keys": split(r.get("related_entity_keys")),
        "evidence_type": r.get("evidence_type"), "note": r.get("note"),
        "correction_status": "human_reviewed"}), "status": status(r.get("review_status")),
        "visibility": "public", "confidence": float(r.get("confidence") or 0.7),
        "created_at": None, "updated_at": None, "data_classification": "tang_changan_curated"
    })

entity_map = {}
for r in entity_rows:
    key = r["entity_key"]
    evidence = [uid("chunk", x) for x in split(r.get("chunk_key"))]
    props = clean({"knowledge_layer": "FACT", "name_traditional": r.get("name_traditional"),
        "category": r.get("category"), "period_start": r.get("period_start"),
        "period_end": r.get("period_end"), "description": r.get("description"),
        "location_text": r.get("location_text"), "source_key": r.get("source_key"),
        "description_quote": r.get("description_quote")})
    if key in entity_map:
        old = entity_map[key]
        old["source_ids"] = sorted(set(old["source_ids"] + evidence))
        old["properties"].update(props)
        old["payload"].update(props)
        old["confidence"] = max(old["confidence"], float(r.get("confidence") or 0.7))
        continue
    entity_map[key] = {
        "id": uid("entity", key), "key": key, "label": r["entity_type"], "name": r["name"],
        "normalized_name": str(r["name"]).casefold().strip(), "aliases": split(r.get("aliases")),
        "properties": props, "source_ids": evidence, "status": status(r.get("review_status")),
        "visibility": "public", "confidence": float(r.get("confidence") or 0.7),
        "payload": {"key": key, "name": r["name"], "label": r["entity_type"], **props},
        "created_at": None, "updated_at": None, "data_classification": "tang_changan_curated"
    }

relation_map = {}
for r in relation_rows:
    key = r["relation_key"]
    evidence = [uid("chunk", x) for x in split(r.get("chunk_key"))]
    props = clean({"source_key": r["source_key"], "target_key": r["target_key"],
        "knowledge_layer": "INTERPRETATION" if r.get("evidence_type") == "\u63a8\u5bfc\u7ed3\u8bba" else "FACT",
        "period_start": r.get("period_start"), "period_end": r.get("period_end"),
        "evidence_quote": r.get("evidence_quote"), "evidence_type": r.get("evidence_type"),
        "note": r.get("note")})
    if key in relation_map:
        old = relation_map[key]
        if old["source_id"] != uid("entity", r["source_key"]) or old["target_id"] != uid("entity", r["target_key"]):
            raise ValueError(f"relation key conflict: {key}")
        old["source_ids"] = sorted(set(old["source_ids"] + evidence))
        quotes = split(old["properties"].get("evidence_quote")) + split(r.get("evidence_quote"))
        old["properties"]["evidence_quote"] = "\n".join(dict.fromkeys(quotes))
        old["payload"]["evidence_quote"] = old["properties"]["evidence_quote"]
        old["confidence"] = max(old["confidence"], float(r.get("confidence") or 0.7))
        continue
    relation_map[key] = {
        "id": uid("relation", key), "key": key, "source_id": uid("entity", r["source_key"]),
        "target_id": uid("entity", r["target_key"]), "type": r["relation_type"],
        "label": r["relation_label"], "properties": props, "source_ids": evidence,
        "status": status(r.get("review_status")), "visibility": "public",
        "confidence": float(r.get("confidence") or 0.7), "payload": {"key": key, **props},
        "created_at": None, "updated_at": None, "data_classification": "tang_changan_curated"
    }

entity_ids = {x["id"] for x in entity_map.values()}
orphans = [r["key"] for r in relation_map.values() if r["source_id"] not in entity_ids or r["target_id"] not in entity_ids]
if orphans:
    raise ValueError(f"orphan relations: {orphans[:10]}")

dataset = {"sources": sources, "source_chunks": chunks,
           "entities": list(entity_map.values()), "relations": list(relation_map.values())}
OUT.mkdir(parents=True, exist_ok=True)
(OUT / "all_in_one.json").write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
(OUT / "validation.json").write_text(json.dumps({"sources": len(sources), "source_chunks": len(chunks),
    "entities": len(entity_map), "relations": len(relation_map), "merged_entity_keys": len(entity_rows)-len(entity_map),
    "merged_relation_keys": len(relation_rows)-len(relation_map), "orphan_relations": len(orphans)}, indent=2), encoding="utf-8")
print(json.dumps({"output": str(OUT / "all_in_one.json"), "sources": len(sources),
    "source_chunks": len(chunks), "entities": len(entity_map), "relations": len(relation_map),
    "merged_relations": len(relation_rows)-len(relation_map)}, ensure_ascii=True))
