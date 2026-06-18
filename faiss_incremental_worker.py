from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

import numpy as np


def normalize_value(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text.replace(".", "", 1).isdigit():
        text = text[:-2]
    return text


def build_faiss_index(vectors: np.ndarray, index_type: str, params: dict) -> object:
    import faiss

    dimension = int(vectors.shape[1])
    normalized_type = normalize_value(index_type or "FlatL2").lower()
    if normalized_type == "hnsw":
        m = int(params.get("hnsw_m", 32))
        index = faiss.IndexHNSWFlat(dimension, m)
        index.hnsw.efSearch = int(params.get("hnsw_ef_search", 64))
    elif normalized_type == "ivf":
        nlist = max(1, int(params.get("ivf_nlist", 64)))
        nlist = min(nlist, max(1, len(vectors)))
        quantizer = faiss.IndexFlatL2(dimension)
        index = faiss.IndexIVFFlat(quantizer, dimension, nlist, faiss.METRIC_L2)
        index.train(vectors)
        index.nprobe = min(max(1, int(params.get("ivf_nprobe", 8))), nlist)
    else:
        index = faiss.IndexFlatL2(dimension)
    index.add(vectors)
    return index


def emit(message: dict) -> None:
    print(json.dumps(message, ensure_ascii=False), flush=True)


def run(input_path: Path, output_path: Path) -> None:
    with np.load(input_path, allow_pickle=False) as data:
        base_vectors = np.asarray(data["base_vectors"], dtype="float32")
        new_vectors = np.asarray(data["new_vectors"], dtype="float32")
        base_labels = [int(value) for value in data["base_labels"].tolist()]
        threshold = float(data["threshold"][0])
        index_type = str(data["index_type"][0])
        params = json.loads(str(data["params"][0]))

    index = build_faiss_index(base_vectors, index_type, params)
    known_labels = list(base_labels)
    next_cluster_id = max(known_labels, default=-1) + 1
    assigned: list[int] = []
    distances: list[float] = []
    total = len(new_vectors)
    for index_position, vector in enumerate(new_vectors):
        query = vector.reshape(1, -1)
        distance, neighbor = index.search(query, 1)
        nearest_distance = float(distance[0][0])
        nearest_index = int(neighbor[0][0])
        if nearest_index >= 0 and nearest_distance <= threshold:
            label = int(known_labels[nearest_index])
        else:
            label = next_cluster_id
            next_cluster_id += 1
        assigned.append(label)
        distances.append(nearest_distance)
        known_labels.append(label)
        index.add(query)
        if index_position + 1 == total or (index_position + 1) % 100 == 0:
            emit({"type": "progress", "done": index_position + 1, "total": total})

    output_path.write_text(
        json.dumps({"assigned": assigned, "distances": distances}, ensure_ascii=False),
        encoding="utf-8",
    )


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("usage: faiss_incremental_worker.py INPUT_NPZ OUTPUT_JSON")
    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    try:
        run(input_path, output_path)
    except Exception as exc:  # noqa: BLE001 - returned to parent process
        emit({"type": "error", "error": str(exc), "traceback": traceback.format_exc(limit=8)})
        raise


if __name__ == "__main__":
    main()
