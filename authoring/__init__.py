"""
authoring — the scenario-editor backend.

A small FastAPI app + a chip<->grammar compiler that lets non-technical authors
build SecretaryBench scenarios visually. It reads and writes the very same
corpus/nodes/*.json files the benchmark consumes (no DB, no export step), and
reuses sb.resolver / sb.schema / sb.oracle for live, real-engine validation.
"""
