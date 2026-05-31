"""SecretaryBench — a temporal-reasoning benchmark for LLM "secretary" agents.

`sb` is the benchmark core: schema/resolver/scheduler/grader/engine drive a
deterministic serve-and-grade loop over a handwritten DAG corpus, validated by
`oracle`. `sb.live` is the networked harness around a real model (runner +
FastAPI store + MCP tool surface). `corpus/` holds the data; `docs/` the design.
"""
