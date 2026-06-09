const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  TabStopType, TabStopPosition, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ─── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  navy:       "1B2A4A",
  darkBlue:   "2E5090",
  midBlue:    "3B6FBF",
  lightBlue:  "D6E4F7",
  veryLight:  "EEF4FC",
  p1:         "4B3F9E",  // purple - RAG
  p2:         "9E3F3F",  // red - Cache
  p3:         "3F7A9E",  // teal - Inventory
  p4:         "9E7A1A",  // gold - Fintech
  p5:         "5A3F9E",  // violet - Agents
  p6:         "9E4A3F",  // brick - Transformer
  p7:         "3F9E6A",  // green - Vector Index
  p8:         "9E3F7A",  // magenta - Quantization
  p9:         "6A9E3F",  // lime - MoE
  p10:        "3F6A9E",  // steel - Speculative
  p11:        "9E6A3F",  // burnt - Eval
  p12:        "3F9E9E",  // cyan - Fine-Tune
  p13:        "7A3F9E",  // indigo - FlashAttn
  p14:        "9E3F4A",  // rose - Offload
  heading:    "1B2A4A",
  accent:     "2E5090",
  muted:      "5A6A8A",
  white:      "FFFFFF",
  offWhite:   "F7F9FC",
  lightGray:  "E8ECF2",
  midGray:    "C0C8D8",
  darkGray:   "4A5568",
  black:      "0D1117",
  warning:    "7A5500",
  warnBg:     "FFF8E8",
  tipBg:      "F0FFF8",
  tip:        "1A6A4A",
};

const CONTENT_W = 9360; // US Letter 1" margins

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const cell = (children, opts = {}) => {
  const cellProps = {
    children,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    borders: opts.borders || {
      top:    { style: BorderStyle.SINGLE, size: 1, color: C.midGray },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: C.midGray },
      left:   { style: BorderStyle.SINGLE, size: 1, color: C.midGray },
      right:  { style: BorderStyle.SINGLE, size: 1, color: C.midGray },
    },
  };

  if (opts.width) {
    cellProps.width = { size: opts.width };
  }
  if (opts.fill) {
    cellProps.shading = { fill: opts.fill };
  }

  return new TableCell(cellProps);
};

const para = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before || 0, after: opts.after || 120, line: opts.line || 276 },
  indent: opts.indent ? { left: opts.indent } : undefined,
  children: [new TextRun({
    text,
    bold: opts.bold || false,
    italics: opts.italic || false,
    color: opts.color || C.black,
    size: opts.size || 22,
    font: "Arial",
  })],
  border: opts.border || undefined,
});

const mixedPara = (runs, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before || 0, after: opts.after || 120, line: 276 },
  indent: opts.indent ? { left: opts.indent } : undefined,
  children: runs.map(r => new TextRun({
    text: r.text,
    bold: r.bold || false,
    italics: r.italic || false,
    color: r.color || C.black,
    size: r.size || 22,
    font: "Arial",
    break: r.break || undefined,
  })),
});

const heading1 = (text, color = C.heading) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 320, after: 160 },
  children: [new TextRun({ text, bold: true, color, size: 36, font: "Arial" })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 4 } },
});

const heading2 = (text, color = C.heading) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 120 },
  children: [new TextRun({ text, bold: true, color, size: 28, font: "Arial" })],
});

const heading3 = (text, color = C.darkBlue) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true, color, size: 24, font: "Arial" })],
});

const sectionBreak = () => new Paragraph({
  spacing: { before: 0, after: 0 },
  children: [new PageBreak()],
});

const bulletItem = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullets", level },
  spacing: { before: 40, after: 60 },
  children: [new TextRun({ text, size: 21, font: "Arial", color: C.black })],
});

const numItem = (text, level = 0) => new Paragraph({
  numbering: { reference: "numbers", level },
  spacing: { before: 40, after: 60 },
  children: [new TextRun({ text, size: 21, font: "Arial", color: C.black })],
});

const boldBullet = (label, text, color = C.darkBlue) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { before: 40, after: 60 },
  children: [
    new TextRun({ text: label + ": ", bold: true, size: 21, font: "Arial", color }),
    new TextRun({ text, size: 21, font: "Arial", color: C.black }),
  ],
});

const spacer = (size = 120) => new Paragraph({ spacing: { before: 0, after: size }, children: [new TextRun("")] });

// Info box table (colored header + body rows)
const infoBox = (title, rows, headerColor, headerBg) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [
    new TableRow({ children: [cell([
      para(title, { bold: true, color: headerColor || C.white, size: 20, after: 0, before: 0 })
    ], { fill: headerBg || C.navy, width: CONTENT_W })] }),
    ...rows.map((r, i) => new TableRow({ children: [cell([
      typeof r === 'string'
        ? para(r, { size: 20, color: C.darkGray, after: 0, before: 0 })
        : r
    ], { fill: i % 2 === 0 ? C.offWhite : C.white, width: CONTENT_W })] }))
  ],
});

// Two-column table
const twoColTable = (leftTitle, rightTitle, leftRows, rightRows, hColor, hBg) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
  rows: [
    new TableRow({ children: [
      cell([para(leftTitle,  { bold: true, color: hColor || C.white, size: 20, after: 0 })], { fill: hBg || C.navy, width: CONTENT_W / 2 }),
      cell([para(rightTitle, { bold: true, color: hColor || C.white, size: 20, after: 0 })], { fill: hBg || C.navy, width: CONTENT_W / 2 }),
    ]}),
    ...Array.from({ length: Math.max(leftRows.length, rightRows.length) }).map((_, i) =>
      new TableRow({ children: [
        cell([para(leftRows[i]  || "", { size: 20, color: C.darkGray, after: 0 })], { fill: C.offWhite, width: CONTENT_W / 2 }),
        cell([para(rightRows[i] || "", { size: 20, color: C.darkGray, after: 0 })], { fill: C.white,    width: CONTENT_W / 2 }),
      ]})
    ),
  ],
});

// Week schedule row table
const weekTable = (weeks) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1200, 2160, 3500, 2500],
  rows: [
    new TableRow({ children: [
      cell([para("Week", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 1200 }),
      cell([para("Focus Area", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 2160 }),
      cell([para("Key Tasks", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 3500 }),
      cell([para("Deliverable", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 2500 }),
    ]}),
    ...weeks.map((w, i) => new TableRow({ children: [
      cell([para(`Wk ${w.num}`, { bold: true, color: C.darkBlue, size: 19, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 1200 }),
      cell([para(w.title, { bold: true, color: C.black, size: 19, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2160 }),
      cell(w.tasks.map(t => para("• " + t, { size: 18, color: C.darkGray, after: 40 })), { fill: i % 2 === 0 ? C.veryLight : C.white, width: 3500 }),
      cell([para(w.deliverable, { size: 18, color: C.muted, italic: true, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2500 }),
    ]}))
  ],
});

// Metrics/KPI table
const metricsTable = (items) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2340, 2340, 2340, 2340],
  rows: [
    new TableRow({ children: ["Metric", "Target", "Method", "Phase"].map((h, i) =>
      cell([para(h, { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.darkBlue, width: 2340 })
    )}),
    ...items.map((row, i) => new TableRow({ children: row.map(t =>
      cell([para(t, { size: 19, color: C.darkGray, after: 0 })], { fill: i % 2 === 0 ? C.offWhite : C.white, width: 2340 })
    )}))
  ],
});

// Warning / tip box
const noteBox = (label, text, color, bg) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [600, CONTENT_W - 600],
  rows: [new TableRow({ children: [
    cell([para(label, { bold: true, color, size: 19, after: 0, align: AlignmentType.CENTER })], { fill: bg, width: 600 }),
    cell([para(text, { size: 19, color, after: 0 })], { fill: bg, width: CONTENT_W - 600 }),
  ]})],
});

// Phase banner
const phaseBanner = (phaseNum, phaseTitle, weeks, description) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [new TableRow({ children: [cell([
    para(`PHASE ${phaseNum}  ·  ${phaseTitle}  ·  ${weeks}`, { bold: true, color: C.white, size: 26, after: 60, before: 40 }),
    para(description, { color: "D0DCF0", size: 20, after: 40, italic: true }),
  ], { fill: C.navy, width: CONTENT_W })] })],
});

// Project header banner
const projectBanner = (id, title, weeks, color) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1440, CONTENT_W - 1440],
  rows: [new TableRow({ children: [
    cell([
      para(id, { bold: true, color: C.white, size: 28, after: 20, align: AlignmentType.CENTER }),
      para(weeks, { color: "D0DCF0", size: 18, after: 0, align: AlignmentType.CENTER }),
    ], { fill: color, width: 1440 }),
    cell([
      para(title, { bold: true, color: C.white, size: 24, after: 40 }),
    ], { fill: C.navy, width: CONTENT_W - 1440 }),
  ]})],
});

// ─── DOCUMENT DATA ────────────────────────────────────────────────────────────
const phases = [
  { num: "I",   title: "High-Performance Backend Infrastructure", weeks: "Weeks 1–12",  desc: "Foundational systems engineering — vector databases, distributed caching, atomic concurrency, and real-time fintech interfaces." },
  { num: "II",  title: "Agentic Orchestration & AI Workspaces",   weeks: "Weeks 13–15", desc: "Multi-model orchestration with role-assigned LLMs, event-driven pipelines, and human-in-the-loop governance." },
  { num: "III", title: "Core Neural Systems Engineering",         weeks: "Weeks 16–25", desc: "From-scratch neural mathematics — transformer architecture, spatial indexing algorithms, and hardware-aware quantization." },
  { num: "IV",  title: "Advanced Inference Scale & MLOps",        weeks: "Weeks 26–34", desc: "Production-grade inference techniques — sparse routing, speculative decoding, automated evaluation, and parameter-efficient fine-tuning." },
  { num: "V",   title: "Hardware & Memory Optimization",          weeks: "Weeks 35–40", desc: "Low-level hardware engineering — GPU memory tiling, PCIe bus profiling, and heterogeneous CPU/VRAM model offloading." },
];

const projects = [

  // ── PROJECT 1 ──────────────────────────────────────────────────────────────
  {
    id: "P1", color: C.p1, phase: "I", weeks: "Weeks 1–4",
    title: "Local AI Evaluator & RAG Optimization Pipeline",
    subtitle: "Vector Mathematics · Open-Source Model Lifecycle · Semantic Query Parsing",
    bottleneck: "Eliminates black-box cloud API token dependency and mitigates semantic drift caused by misaligned chunk boundaries. Traditional cloud-dependent RAG pipelines expose applications to latency spikes, rate-limit failures, and opaque model versioning. This project implements a fully local document ingestion engine using recursive character chunking to map raw text arrays into dense numerical vector spaces, enabling deterministic, cost-free semantic retrieval.",
    overview: "The Local AI Evaluator is a production-quality Retrieval-Augmented Generation (RAG) pipeline designed to operate entirely on local hardware. It ingests raw documents, converts text into dense vector embeddings using open-source models, stores them in a pgvector-enabled PostgreSQL instance, and benchmarks two competing local LLMs against a standardized set of evaluation questions. The system eliminates cloud dependency, making it suitable for air-gapped environments, cost-sensitive deployments, and environments where data sovereignty is mandatory.",
    objectives: [
      "Engineer a document ingestion and recursive chunking pipeline supporting PDF and plain-text inputs.",
      "Configure a native PostgreSQL instance with the pgvector extension for high-dimensional vector storage.",
      "Construct a benchmarking harness evaluating two open-source models against 20 targeted semantic questions.",
      "Develop a Next.js analytics dashboard for real-time performance tracking across multiple evaluation dimensions.",
      "Experimentally validate the impact of chunk size and overlap parameters on retrieval quality.",
    ],
    stack: {
      "Language & Runtime": "Python 3.11, Node.js 20 (LTS)",
      "AI / ML Framework": "LangChain, Ollama REST API",
      "Models": "llama3.2:3b (fast inference), deepseek-r1:7b (heavy reasoning), nomic-embed-text (embeddings)",
      "Database": "PostgreSQL 16 + pgvector extension",
      "Frontend": "Next.js 14 (App Router), Recharts, Tailwind CSS",
      "Infrastructure": "Docker Desktop, Docker Compose",
      "Testing": "Manual accuracy scoring (1–5 Likert scale), automated timing via Python time module",
    },
    architecture: [
      "Document Ingestion Layer: LangChain PDF loaders parse raw files, feeding text to the recursive splitter. Chunk metadata (source file, page number, chunk index) is preserved as JSON alongside each text segment.",
      "Embedding Generation Layer: Text chunks are serialized and dispatched to the Ollama REST endpoint (/api/embeddings). The nomic-embed-text model returns 768-dimensional float32 vectors. Chunks are batch-inserted into PostgreSQL using parameterized queries with psycopg2.",
      "Vector Retrieval Engine: At query time, the question is independently embedded and a cosine similarity search (pgvector <=> operator) retrieves the top-K most semantically relevant chunks. K is configurable; recommended starting value is 3.",
      "Benchmarking Harness: The benchmark script loops over 20 predefined evaluation questions, constructs a context-augmented prompt (retrieved chunks + question), dispatches to both models via concurrent HTTP calls, and logs response latency, generation speed, and manual accuracy scores to a structured JSON file.",
      "Analytics Dashboard: A Next.js frontend consumes the benchmark JSON via an API route and renders three interactive Recharts visualizations. A live test input field allows ad-hoc question evaluation with results appended dynamically to all charts.",
    ],
    weekPlan: [
      { num: 1, title: "Ingestion & Chunking", tasks: ["Install Ollama, pull llama3.2:3b, deepseek-r1:7b, nomic-embed-text", "Create monorepo project structure and Docker Compose base", "Write Python ingestion script using LangChain document loaders", "Implement RecursiveCharacterTextSplitter with variable chunk_size (256/512/1024) and overlap", "Log chunk statistics: count, avg token density, source distribution", "Pin all dependencies in requirements.txt on Day 1"], deliverable: "Working ingestion script outputting a structured JSON array of text chunks with metadata" },
      { num: 2, title: "Vector Layer", tasks: ["Launch PostgreSQL + pgvector via Docker Compose", "Create DB schema: documents table, embeddings table with vector(768) column", "Write embedding generation loop — Ollama REST API, batch size 50 to avoid OOM", "Batch-insert chunks and vectors into PostgreSQL with psycopg2", "Validate: manual cosine similarity query returning top-5 most relevant chunks", "Document schema with column-level comments"], deliverable: "PostgreSQL DB populated with embedded chunks; manual similarity query returning coherent results" },
      { num: 3, title: "Benchmark Core", tasks: ["Compose 20 evaluation questions spanning easy/medium/hard difficulty", "Build benchmark loop: question → similarity query → top-3 chunks → context prompt", "Send augmented prompt to both models simultaneously via concurrent HTTP calls", "Log per-model metadata: time-to-first-token, total tokens/sec, full response text", "Add manual accuracy scorer (1–5 Likert) with results appended to benchmark_results.json", "Implement retry logic: 3 retries with 5-second exponential backoff on Ollama timeouts"], deliverable: "benchmark_results.json containing complete metadata for all 20 questions × 2 models" },
      { num: 4, title: "Analytics Dashboard", tasks: ["Scaffold Next.js 14 project with App Router, Tailwind CSS, dynamic Recharts imports", "Build API route /api/results serving benchmark_results.json to the frontend", "Render latency comparison bar chart (model A vs. B per question)", "Render tokens-per-second line graph over benchmark run sequence", "Render accuracy scatter plot: manual score vs. latency per model", "Add live 'Run Test' input: new question fires benchmark, appends to all charts in real time"], deliverable: "Next.js dashboard with three interactive charts and a live question-testing input field" },
      { num: 5, title: "RAG Quality Tuning", tasks: ["Experiment with chunk sizes (256/512/1024/2048) — re-run benchmark, record accuracy delta", "Implement MMR (Maximal Marginal Relevance) retrieval to reduce chunk redundancy", "Add metadata-filtered retrieval: constrain similarity search by document source", "Build comparison grid in dashboard: chunk_size × model × average accuracy score", "Refactor Python codebase into clean module structure (ingest.py, embed.py, benchmark.py)", "Add .env configuration for all model names and database connection strings"], deliverable: "Tuned pipeline with measurable accuracy improvement over Week 3 baseline; refactored codebase" },
      { num: 6, title: "Documentation & Portfolio Polish", tasks: ["Write README.md: architecture overview, setup guide, benchmark result summary table", "Create architecture diagram (Mermaid or draw.io) of full RAG pipeline flow", "Configure Docker Compose one-command startup: docker compose up --build", "Record 2-minute Loom demo: ingest document, run benchmark, navigate dashboard", "Deploy Next.js dashboard to Vercel free tier", "Write reflection note: three things learned, two things to do differently"], deliverable: "Public GitHub repo with README, live Vercel URL, embedded architecture diagram, and Loom demo link" },
    ],
    learnings: [
      "Embedding vector mechanics: understanding cosine similarity, L2 distance, and inner product distance operators in pgvector.",
      "Chunking strategy tradeoffs: how chunk size and overlap affect retrieval precision, recall, and context quality.",
      "Complete RAG loop mechanics: the retrieve-augment-generate pipeline from text input to grounded model output.",
      "LLM performance profiling: measuring and interpreting Time-to-First-Token (TTFT), tokens-per-second, and generation consistency.",
      "Local LLM operations: model lifecycle management with Ollama, API integration patterns, and connection resilience.",
      "Vector database engineering: pgvector indexing strategies, HNSW vs. IVFFlat index selection, and query optimization.",
    ],
    metrics: [
      ["Retrieval Accuracy",  "≥ 70% score on 20 questions", "Manual 1–5 Likert scoring", "Week 3"],
      ["TTFT (llama3.2:3b)",  "< 3 seconds per query",       "Python time module",        "Week 3"],
      ["Chunk Improvement",   "≥ 15% accuracy delta",        "Baseline vs. tuned config", "Week 5"],
      ["Dashboard Load Time", "< 2 seconds initial render",  "Browser DevTools",          "Week 4"],
    ],
    warning: "LangChain versions change rapidly. Pin your requirements.txt on Day 1 or risk breaking API changes mid-project. Use nomic-embed-text for 768-dim vectors — do not configure pgvector columns as vector(1536) (OpenAI-sized); this will cause silent dimension mismatch errors.",
    tip: "Use your own real documents — RunTrack session notes, GastroHub data schemas, or work process documentation. Real, domain-specific data produces benchmark results that are personally meaningful and directly demonstrate the pipeline's practical value.",
  },

  // ── PROJECT 2 ──────────────────────────────────────────────────────────────
  {
    id: "P2", color: C.p2, phase: "I", weeks: "Weeks 5–8",
    title: "Distributed Cache-Aside System with Cache Invalidation",
    subtitle: "In-Memory Datastores · Event-Driven Architecture · Dirty Read Prevention",
    bottleneck: "Prevents database CPU starvation and disk-bound I/O thrashing under heavy virtual user read scaling. Relational databases under high-frequency read workloads reach a saturation point where the disk-bound query executor becomes the system's critical bottleneck. The Cache-Aside pattern inserts a high-speed in-memory layer between the application and the database, dramatically reducing primary storage pressure while introducing cache coherence as a new engineering constraint requiring careful management.",
    overview: "This project implements a production-quality two-tier data architecture combining PostgreSQL as the persistent source of truth and Redis as a high-speed volatile cache. The system enforces the Cache-Aside (lazy loading) pattern with atomic write-path invalidation — ensuring that every cache entry is either fresh or absent, never stale. The project culminates in a rigorous k6 load test that empirically demonstrates the CPU savings delivered by the caching layer under 3,000 concurrent virtual users.",
    objectives: [
      "Deploy a dual-tier PostgreSQL and Redis data layer using Docker Compose.",
      "Implement Cache-Aside read middleware with hit/miss branching, TTL management, and response header telemetry.",
      "Engineer atomic write paths that wrap SQL mutations and Redis cache evictions in a single coherent transaction sequence.",
      "Protect against the cache stampede failure mode using Redis SETNX-based distributed locking.",
      "Conduct k6 load tests under 3,000 virtual users to empirically prove database CPU reduction.",
    ],
    stack: {
      "Language & Runtime": "Node.js 20 (TypeScript), optional Go 1.22 for high-throughput variant",
      "Cache Layer": "Redis 7.2 (ioredis client)",
      "Database": "PostgreSQL 16 (Prisma ORM or raw pg driver)",
      "Load Testing": "k6 v0.51+",
      "Logging": "Pino structured JSON logger",
      "Infrastructure": "Docker Desktop, Docker Compose",
      "Testing": "ioredis-mock for unit tests, custom integration test suite",
    },
    architecture: [
      "Read Path (Cache-Aside): On GET requests, the middleware generates a canonical Redis key (e.g., cache:news:{id}). It issues a Redis GET — if a value is returned (cache hit), the JSON is deserialized and returned with an X-Cache: HIT header. On a cache miss, the middleware fetches from PostgreSQL, serializes the result to JSON, writes it to Redis with a configured TTL, and returns the response with X-Cache: MISS.",
      "Write Path (Atomic Invalidation): On PUT and PATCH requests, the handler opens a PostgreSQL transaction, applies the mutation, and on successful commit fires a Redis DEL command for the affected key. This sequence ensures no stale entry can be read between the write and the invalidation.",
      "Cache Stampede Protection: A SETNX-based mutex lock is applied per cache key during miss resolution. If multiple concurrent requests miss the same key, only one proceeds to query PostgreSQL; others wait briefly on the lock and then read the now-populated cache entry.",
      "Observability Layer: A /metrics endpoint tracks hit rate, miss rate, average response latency, and total request count. Redis INCR commands on telemetry keys provide non-blocking stat collection.",
      "TTL Strategy: Cache entries expire automatically after a configured TTL (default 60 seconds). A separate cache warming script pre-populates the top 1,000 most-accessed records on server startup, eliminating cold-start miss spikes.",
    ],
    weekPlan: [
      { num: 5, title: "Two-Tier Database Setup", tasks: ["Docker Compose: PostgreSQL + Redis containers with named volumes", "Design schema: news_items table with 10+ columns including full-text content", "Write faker.js seeder: generate and insert 20,000 realistic records", "Build baseline Express/TypeScript CRUD endpoints — zero caching", "Measure baseline GET latency: log avg, p95, p99 response times", "Configure Pino structured JSON logging middleware for all endpoints"], deliverable: "Express API with 20,000 seeded records and documented baseline latency measurements" },
      { num: 6, title: "Cache-Aside Middleware", tasks: ["Install ioredis, configure Redis connection with exponential reconnect strategy", "Implement Cache-Aside GET: Redis check → hit return → miss: PG fetch → Redis SET with TTL", "Add cache hit/miss counter via Redis INCR on telemetry keys", "Set X-Cache: HIT/MISS response header on all GET endpoints", "Manual verification: call GET twice for same ID; confirm second response < 2ms", "Experiment with TTL values: 15s vs. 60s vs. 300s — measure hit rate change"], deliverable: "Cache-aside GET endpoint with sub-2ms cache hits and observable hit/miss ratio metrics" },
      { num: 7, title: "Atomic Cache Invalidation", tasks: ["Build PUT endpoint: SQL UPDATE wrapped in transaction + Redis DEL on commit", "Build PATCH endpoint with partial field updates and identical invalidation", "Write dirty-read verification test: update → immediate GET → assert fresh data returned", "Build DELETE endpoint: SQL delete + Redis key cleanup", "Implement SETNX mutex lock per cache key to prevent cache stampede", "Write integration test suite: create → cache read → update → fresh read → delete"], deliverable: "Full CRUD with atomic invalidation; dirty-read test passing; stampede protection operational" },
      { num: 8, title: "k6 Load Test & Performance Proof", tasks: ["Write k6 smoke test: 10 VUs × 30 seconds on GET endpoint", "Write ramp-up load test: 0 → 3,000 VUs over 60s, hold 30s, ramp down", "Run load test with caching DISABLED — capture PG CPU%, RPS, p95 latency", "Re-run identical test with caching ENABLED — capture identical metrics", "Export k6 JSON results and render comparison chart", "Screenshot side-by-side DB CPU comparison — primary portfolio proof artifact"], deliverable: "k6 results JSON + comparison chart proving ≥70% database CPU reduction under 3,000 VU load" },
      { num: 9, title: "Monitoring & Observability", tasks: ["Build /metrics endpoint: hit rate, miss rate, avg latency, total requests", "Build admin dashboard page displaying real-time cache statistics", "Implement cache warming script: pre-populate top 1,000 records on startup", "Write unit tests for cache middleware using ioredis-mock", "Document TTL strategy decision with measured hit-rate data", "Create sequence diagram of full Cache-Aside read/write flow"], deliverable: "Metrics endpoint, admin dashboard, unit test suite, and cache warming implementation" },
      { num: 10, title: "Documentation & Portfolio Polish", tasks: ["Write README: Cache-Aside pattern explanation, architecture diagram, setup guide", "Write technical blog post: 'How Cache-Aside Shielded My Database at 3,000 RPS'", "Deploy to Railway free tier (Node.js + PostgreSQL + Redis as one project)", "Record Loom: run k6 test live, narrate the CPU comparison graphs", "Embed k6 load test screenshots in GitHub README", "Publish blog post to Dev.to or Hashnode"], deliverable: "Public GitHub repo, Railway deployment URL, technical blog post, and Loom demo recording" },
    ],
    learnings: [
      "Cache-Aside topology and its tradeoffs versus Write-Through, Read-Through, and Write-Behind caching strategies.",
      "TTL management: the precision tradeoff between freshness guarantees and cache hit rate efficiency.",
      "Isolation of dirty reads through transactional pipeline commits — ensuring atomicity across heterogeneous storage layers.",
      "Cache stampede failure mode: identification, simulation, and prevention via distributed locking with Redis SETNX.",
      "k6 load testing methodology: virtual user ramping, threshold configuration, and metric interpretation.",
      "Structured observability design: building metrics endpoints and dashboards alongside the application rather than as an afterthought.",
    ],
    metrics: [
      ["Cache Hit Rate",       "≥ 85% under steady load",   "Redis INCR telemetry counters",  "Week 6"],
      ["Cache Hit Latency",    "< 2ms P99",                  "k6 response time metrics",       "Week 6"],
      ["DB CPU Reduction",     "≥ 70% vs. uncached baseline","Docker stats during k6 run",     "Week 8"],
      ["Dirty Read Prevention","0 stale reads in test suite","Integration test assertions",    "Week 7"],
    ],
    warning: "Cache stampede is the most dangerous production failure mode in this architecture. When 100 concurrent requests miss the same key simultaneously, they all query PostgreSQL — multiplying database load exactly when you are trying to reduce it. Implement SETNX locking before running any load test, not after.",
    tip: "Your Power Automate and MES background is directly relevant here. Production floor data refresh intervals in MES systems face the exact same TTL design problem — too short and you hammer the database, too long and operators see stale machine states. The mental model transfers perfectly.",
  },

  // ── PROJECT 3 ──────────────────────────────────────────────────────────────
  {
    id: "P3", color: C.p3, phase: "I", weeks: "Weeks 9–12",
    title: "High-Performance Inventory Reservation Engine",
    subtitle: "Thread Safety · Atomic Lua Scripts · Eventual Consistency Architecture",
    bottleneck: "Eliminates race conditions, negative data balances, and transaction locking delays in long-running database connection pools. In high-concurrency checkout systems, the check-and-decrement operation — read stock, evaluate condition, write decrement — is non-atomic when implemented as separate operations. Concurrent threads interleave these steps, producing negative inventory counts and overselling. This project replaces the unsafe three-step pattern with a Redis Lua script that executes atomically on Redis's single-threaded engine.",
    overview: "The Inventory Reservation Engine solves one of the most consequential concurrency problems in e-commerce and booking systems: overselling. The system exposes a high-throughput checkout endpoint backed by Redis in-memory stock counters. A custom Lua script performs atomic check-and-decrement, making it physically impossible for two concurrent requests to both claim the last unit. Successful reservations are enqueued via BullMQ and persisted to PostgreSQL by an asynchronous worker pool, decoupling the latency-sensitive checkout path from the I/O-bound database write.",
    objectives: [
      "Build an intentionally broken non-atomic checkout endpoint to empirically demonstrate the race condition.",
      "Write a Redis Lua script that atomically checks stock availability and decrements the counter in a single uninterruptible operation.",
      "Deploy a BullMQ asynchronous worker pipeline to handle background PostgreSQL persistence.",
      "Execute a chaos test: 2,000 concurrent requests against exactly 5 available units — asserting precisely 5 successes.",
      "Implement circuit breaker and dead-letter queue patterns for production resilience.",
    ],
    stack: {
      "Language & Runtime": "Node.js 20 (TypeScript)",
      "Atomic Layer": "Redis 7.2 + custom Lua scripts (EVALSHA)",
      "Queue System": "BullMQ v5 (Redis Streams-based)",
      "Database": "PostgreSQL 16",
      "Concurrency Testing": "Node.js Promise.all with axios, custom assertion suite",
      "Monitoring": "BullMQ Board UI, Prometheus metrics endpoint",
      "Infrastructure": "Docker Compose (reuses P2 base configuration)",
    },
    architecture: [
      "Checkout Gateway: The endpoint extracts product ID and user ID from the request. It calls the Redis Lua script via EVALSHA. The script is pre-loaded into Redis using SCRIPT LOAD, and its SHA1 hash is cached for subsequent calls, eliminating re-transmission overhead.",
      "Lua Atomic Script: The script executes as follows — (1) retrieve current stock value, (2) if stock ≤ 0, return failure code 0, (3) decrement stock by 1 using DECRBY, (4) return success code 1. Because Redis processes Lua scripts synchronously on a single thread, no external operation can interrupt between steps 2 and 3.",
      "Reservation Expiry: Accepted reservations carry a 10-minute expiry TTL. If payment confirmation does not arrive within the window, a separate expiry worker restores the stock counter, preventing ghost reservations from permanently removing inventory.",
      "Async Worker Pipeline: On successful checkout, the endpoint pushes an order event (userId, productId, quantity, timestamp, reservationId) to a BullMQ queue. The worker pool consumes events and writes invoice records to PostgreSQL. Failed writes retry three times with exponential backoff before being routed to a dead-letter queue.",
      "Chaos Assertion Suite: The test script dispatches 2,000 concurrent checkout requests using Promise.all with axios. Post-execution, it asserts four invariants: exactly 5 HTTP 200 responses, exactly 1,995 HTTP 409 rejections, Redis stock counter equals 0, and PostgreSQL invoice count equals 5.",
    ],
    weekPlan: [
      { num: 9,  title: "Concurrency Gateway (Broken First)", tasks: ["Docker Compose: Redis + PostgreSQL (extend P2 Compose file)", "Seed Redis with product stock counters: product:stock:{id} = 100 for 10 products", "Build checkout endpoint WITHOUT atomic protection (read → check → decrement as three steps)", "Run 500 concurrent axios requests to buy the same product — observe race condition", "Document the bug: screenshot negative stock, log overselling events", "Draw timeline diagram showing two threads interleaving the three non-atomic steps"], deliverable: "Broken checkout endpoint with documented race condition proof — screenshot and timeline diagram" },
      { num: 10, title: "The Lua Atomic Lock", tasks: ["Write Redis Lua script: IF stock > 0 THEN DECRBY 1 RETURN 1 ELSE RETURN 0", "Load script via SCRIPT LOAD, cache SHA1 hash, call via EVALSHA in checkout handler", "Re-run 500-concurrent-request test — assert zero overselling, exact stock count", "Add 10-minute reservation expiry: EXPIRE key on reservation, restore on timeout", "Write unit tests for Lua script using ioredis-mock with custom Lua evaluator", "Load test with 1,000 concurrent requests — verify deterministic results across 3 runs"], deliverable: "Atomic Lua checkout with zero race conditions and deterministic stock accounting across all test runs" },
      { num: 11, title: "The Event Pipeline", tasks: ["Install BullMQ v5, configure order-processing queue backed by Redis", "Push order event payload to queue on successful checkout: {userId, productId, qty, timestamp}", "Write worker process: consume events, write invoice records to PostgreSQL with idempotency guard", "Add retry logic: 3 retries with exponential backoff (1s, 4s, 16s) for failed DB writes", "Route permanently failed jobs to dead-letter queue: 'orders:failed' queue with full error metadata", "Deploy BullMQ Board UI for real-time queue visualization: active, waiting, completed, failed counts"], deliverable: "Complete async pipeline: checkout → BullMQ queue → worker → PostgreSQL with dead-letter handling" },
      { num: 12, title: "The Chaos Script", tasks: ["Write chaos test script: 2,000 concurrent axios requests, product with exactly 5 units", "Assert Invariant 1: exactly 5 HTTP 200 responses in results array", "Assert Invariant 2: exactly 1,995 HTTP 409 rejection responses", "Assert Invariant 3: Redis stock counter === 0 after all requests complete", "Assert Invariant 4: PostgreSQL invoice count === 5 with no duplicate records", "Run chaos test 3 times — confirm identical, deterministic results across all runs"], deliverable: "Chaos test script with 4 passing invariant assertions across 3 deterministic runs — screenshot proof" },
      { num: 13, title: "Resilience Hardening", tasks: ["Add Redis connection pooling to prevent connection exhaustion under high throughput", "Implement basic circuit breaker: if Redis unavailable, fall back to DB-only pessimistic lock", "Build stock replenishment admin endpoint with full PostgreSQL audit trail", "Stress test: 5,000 RPS sustained for 30 seconds — measure queue depth and worker lag", "Tune BullMQ worker concurrency: test 1, 5, and 10 workers — find optimal throughput", "Add Prometheus metrics endpoint: checkout_success_total, checkout_fail_total, queue_depth"], deliverable: "Hardened system with circuit breaker, metrics endpoint, and stress test results documentation" },
      { num: 14, title: "Documentation & Portfolio Polish", tasks: ["Write README: Lua atomicity explanation, BullMQ pipeline diagram, chaos test results table", "Create system architecture diagram: client → API → Lua/Redis → BullMQ → PostgreSQL worker", "Write 'before vs. after' comparison: race condition screenshot vs. chaos test assertion results", "Deploy to Railway: API + Redis + PostgreSQL + BullMQ worker as separate Railway services", "Record Loom: run chaos script live, narrate 5 successes and 1,995 failures, explain invariants", "Write post-mortem doc: remaining edge cases (Redis restart, partial worker failures, split-brain)"], deliverable: "Public GitHub repo, Railway deployment, Loom demo, post-mortem document" },
    ],
    learnings: [
      "Atomic transactional steps in single-threaded execution loops: understanding why Redis's single-threaded model makes Lua scripts safe.",
      "Race condition identification and empirical proof: building the broken version first to demonstrate the problem before solving it.",
      "Decoupled persistent I/O via asynchronous queue design: the checkout path never touches PostgreSQL directly.",
      "Eventual consistency tradeoffs: the system is fast because it accepts that the database lags behind Redis by milliseconds.",
      "Dead-letter queue design: handling permanent failures without data loss and preserving observability into failure modes.",
      "Chaos engineering methodology: designing tests with explicit invariant assertions rather than simply checking for absence of errors.",
    ],
    metrics: [
      ["Overselling Rate",      "0 units (absolute)",         "Chaos test Invariant 1 + 4",   "Week 12"],
      ["Checkout Latency",      "< 10ms P99 (Redis gate)",    "Custom timing middleware",      "Week 10"],
      ["Worker Throughput",     "≥ 1,000 orders/min",         "BullMQ Board completed count", "Week 11"],
      ["Queue Error Rate",      "< 0.1% after retries",       "Dead-letter queue depth",      "Week 11"],
    ],
    warning: "Lua scripts in Redis cannot invoke external services, perform async I/O, or call other Redis connections. Keep the script purely computational: read, evaluate, write, return. Any attempt to call an external HTTP endpoint or use Lua coroutines will result in a Redis error and a broken checkout.",
    tip: "This is your single strongest portfolio artifact across all 14 projects. The Loom recording of the chaos test — showing 5 successes and 1,995 clean rejections — is the kind of empirical engineering proof that makes senior developers pay attention in technical interviews.",
  },

  // ── PROJECT 4 ──────────────────────────────────────────────────────────────
  {
    id: "P4", color: C.p4, phase: "I", weeks: "Weeks 1–12 (Parallel)",
    title: "Arveend Fintech Terminal",
    subtitle: "Real-Time Financial Intelligence · WebSocket Streaming · Paper Trading Simulation",
    bottleneck: "Resolves unsynchronized UI states, client-side DOM re-render spikes, and external API rate limit failures. Traditional REST-polling financial UIs generate excessive API calls, produce visible latency jags between data refresh cycles, and suffer from full-component re-renders on every tick update. This project replaces polling with persistent WebSocket tunnels and implements a smart caching proxy to buffer external API traffic within free-tier rate limits.",
    overview: "The Arveend Terminal is a full-stack financial intelligence platform designed to serve as the capstone integration of Projects 1 through 3. It aggregates real-time market data via third-party APIs, processes it through local AI sentiment and RAG pipelines from P1, caches profile data using the Cache-Aside pattern from P2, and enforces atomic paper trading order execution using the reservation engine pattern from P3. The system is built in three phases: a core REST analytics engine (V1), a local AI intelligence layer (V2), and a reactive WebSocket streaming interface (V3).",
    objectives: [
      "Build a V1 core REST engine: stock search, company profiles, candlestick charts, and portfolio P&L tracker.",
      "Build a V2 local intelligence layer: Ollama-powered news sentiment scoring, RAG earnings summarizer, risk metric calculations, and a multi-column market screener.",
      "Build a V3 reactive streaming interface: WebSocket price feeds, in-memory alert processing, and a paper trading simulation wallet.",
      "Mock all external API responses with local JSON fixtures to develop independently of rate limit constraints.",
      "Deploy V1 to a public URL by the end of Week 4 to establish a shareable portfolio link early.",
    ],
    stack: {
      "Frontend": "Next.js 14 (App Router), Recharts / lightweight-charts, Tailwind CSS",
      "Backend API": "Node.js 20 (TypeScript), Express or Fastify",
      "Real-Time": "Socket.io (WebSocket server + client)",
      "Cache Layer": "Redis 7.2 (Cache-Aside pattern from P2)",
      "Database": "PostgreSQL 16 (watchlist, portfolio, transactions, alerts)",
      "AI Layer": "Ollama (llama3.2:3b for sentiment), pgvector (RAG from P1 pipeline)",
      "External Data": "Finnhub API (free tier, ~60 calls/min)",
      "Auth": "JWT (jsonwebtoken), bcrypt for password hashing",
      "Deployment": "Railway or Fly.io (free tier)",
    },
    architecture: [
      "V1 — Data Aggregation Layer: The backend proxies Finnhub API calls through a caching middleware. Company profile data is cached in Redis with a 24-hour TTL. Historical OHLCV price data is stored in PostgreSQL to avoid re-fetching. All API responses are mirrored to local JSON fixture files on first fetch, enabling offline development.",
      "V1 — Portfolio Engine: The PostgreSQL schema stores watchlists (user_id, symbol, added_at), transactions (user_id, symbol, side, quantity, price, executed_at), and computed P&L views. A SQL aggregation function computes total portfolio value, unrealised gain/loss, and percentage return dynamically against live quotes.",
      "V2 — Sentiment Pipeline: An async service fetches the 10 latest news articles for a given symbol from Finnhub. Article headlines and summaries are concatenated into a structured prompt and dispatched to llama3.2:3b via the Ollama REST API. The model returns a JSON sentiment object: {score: 0.75, label: 'Bullish', reasoning: '...'}.",
      "V2 — RAG Earnings Summarizer: Company earnings call transcripts are stored as chunks in the pgvector database from P1. When a user clicks 'AI Summary', the system performs a similarity search against the earnings corpus and synthesises the top-5 retrieved chunks into a structured highlights-and-risks report.",
      "V3 — WebSocket Tick Feed: A Socket.io server emits mock price tick events every 500ms per subscribed symbol. Each tick carries {symbol, price, change, volume, timestamp}. The client-side chart module uses a circular buffer to append incoming ticks to the visible window without triggering a full React re-render.",
      "V3 — Paper Trading Gate: Buy and sell order execution uses the Redis Lua atomic pattern from P3. The virtual cash balance and position quantities are maintained as Redis counters. Successful trades are persisted to PostgreSQL by a BullMQ worker, mirroring the exact inventory reservation architecture.",
    ],
    weekPlan: [
      { num: 1,  title: "API Aggregation + Mock Layer", tasks: ["Register Finnhub free API key", "Build proxy endpoints: stock search, company profile, OHLCV historical data, RSS news feed", "Implement mock fixture layer: mirror all Finnhub responses to local JSON on first call", "Configure Redis cache for profile endpoints with 24-hour TTL", "Write OpenAPI schema for all backend routes"], deliverable: "Backend API serving stock data with Redis cache and JSON fixture fallback layer" },
      { num: 2,  title: "Database Schema + Auth", tasks: ["Design PostgreSQL schema: users, watchlists, transactions, portfolio_snapshots, alerts", "Implement JWT authentication: register, login, token refresh endpoints", "Write SQL P&L aggregation function: compute gain/loss against live quotes", "Seed database with 5 sample portfolios for development testing"], deliverable: "Authenticated API with portfolio schema, P&L computation, and JWT auth middleware" },
      { num: 3,  title: "Visual Layouts (V1 Frontend)", tasks: ["Scaffold Next.js 14 with App Router, Tailwind, dynamic Recharts imports", "Build symbol search bar with debounced autocomplete", "Render candlestick chart for historical OHLCV data using lightweight-charts", "Build portfolio dashboard: holdings table with live P&L, gain/loss color coding"], deliverable: "Functional V1 frontend with candlestick charts and portfolio P&L dashboard" },
      { num: 4,  title: "V1 Deploy + State Management", tasks: ["Connect frontend to backend via React Query for server state management", "Implement loading states, error boundaries, and empty-state components", "Configure HTTP cache headers on profile endpoints (Cache-Control: max-age=300)", "Deploy V1 to Railway/Fly.io — establish live public URL", "Write end-to-end smoke test: search → profile → chart → portfolio view"], deliverable: "Live V1 deployment at a public URL — shareable portfolio link established" },
      { num: 5,  title: "V2 AI Sentiment Parser", tasks: ["Integrate Ollama llama3.2:3b into backend via REST API client", "Build async sentiment service: fetch 10 articles → batch prompt → JSON sentiment response", "Display sentiment score on company profile page: percentage + bull/bear label + reasoning", "Add sentiment history chart: track score over 7-day rolling window"], deliverable: "Live AI-powered sentiment scoring displayed on company profile pages" },
      { num: 6,  title: "V2 RAG Earnings Summarizer", tasks: ["Import pgvector connection and embedding functions from P1 pipeline", "Store earnings call transcripts as chunks in pgvector database", "Build 'AI Summary' button on company page — triggers RAG query and displays highlights/risks", "Add source citation: show which transcript chunk grounded each summary point"], deliverable: "One-click AI earnings summarizer powered by P1 RAG pipeline with source citations" },
      { num: 7,  title: "V2 Risk Metrics + Screener", tasks: ["Write backend math: beta calculation, 20/50-day moving average crossover detection, dividend yield", "Build screener data table: 50 symbols with all computed metrics pre-cached in Redis", "Implement multi-column filter UI: PE ratio, dividend yield, beta, market cap range sliders", "Add CSV export of screener results"], deliverable: "Market screener with multi-column filters over 50 pre-computed symbol metrics" },
      { num: 8,  title: "V2 Polish + Integration", tasks: ["Connect all V2 features into unified company profile page layout", "Add loading skeletons for async AI operations (sentiment + RAG summary)", "Performance audit: profile React renders, eliminate unnecessary re-renders with useMemo", "Write integration tests for sentiment pipeline and screener filter logic"], deliverable: "Fully integrated V2 feature set with performance-optimised company profile page" },
      { num: 9,  title: "V3 WebSocket Server", tasks: ["Set up Socket.io server alongside Express API on separate port", "Implement symbol subscription rooms: clients join/leave rooms per symbol", "Emit mock price ticks every 500ms per subscribed room: {symbol, price, change, volume, ts}", "Add reconnection logic with exponential backoff on client side"], deliverable: "WebSocket server emitting price ticks to subscribed clients at 500ms intervals" },
      { num: 10, title: "V3 Client Sync", tasks: ["Build useWebSocket hook consuming Socket.io events in Next.js", "Implement circular buffer for tick history: maintain last 200 ticks per symbol in component state", "Append incoming ticks to chart without full re-render using imperative chart API", "Add connection status indicator: connected/reconnecting/disconnected badge on UI"], deliverable: "Real-time chart updating smoothly with WebSocket ticks and no full re-renders" },
      { num: 11, title: "V3 Alerts + Paper Trading", tasks: ["Build in-memory alert processor in Redis: store target price per symbol per user", "On each tick, evaluate alert conditions: if price crosses threshold, emit socket alert event", "Implement paper trading wallet: virtual cash balance + position tracking in Redis (Lua atomic pattern from P3)", "Build trade execution UI: market buy/sell order form with confirmation step and position P&L"], deliverable: "Alert notification system and paper trading wallet with atomic order execution" },
      { num: 12, title: "System Consolidation", tasks: ["Code cleanup: consistent TypeScript types, remove all console.log debug statements", "Write integration tests for WebSocket channels: connect, subscribe, tick, unsubscribe", "Performance audit: measure and document end-to-end latency from tick emission to chart render", "Draft GitHub README: architecture diagram, feature list, tech stack, live demo link", "Final deploy: push all V3 features to Railway production environment"], deliverable: "Production-ready Arveend Terminal with full test coverage, documentation, and live deployment" },
    ],
    learnings: [
      "Low-latency WebSocket environment management: room subscription patterns, client reconnection strategies, and server-side event broadcasting.",
      "Reactive UI optimisation under heavy data feeds: circular buffer state management, imperative chart API usage, and useMemo/useCallback optimisation.",
      "Transaction history logic and portfolio accounting: cost basis calculation, unrealised P&L computation, and position reconciliation.",
      "Third-party API rate limit management: fixture-layer development, intelligent caching strategies, and graceful degradation.",
      "Full-stack integration of Projects 1–3: applying RAG, Cache-Aside, and atomic reservation patterns within a single coherent product.",
      "Production deployment pipeline: environment variable management, Railway service configuration, and end-to-end smoke testing.",
    ],
    metrics: [
      ["Tick-to-Chart Latency",  "< 100ms end-to-end",         "Browser DevTools timeline",    "Week 10"],
      ["P&L Computation",        "< 200ms for 50-stock portfolio","PostgreSQL EXPLAIN ANALYZE",  "Week 2"],
      ["Sentiment Response Time","< 30s (local Ollama)",        "Backend API timing middleware", "Week 5"],
      ["V1 Deploy",              "Live URL by end of Week 4",   "Railway deployment dashboard", "Week 4"],
    ],
    warning: "Finnhub's free tier allows approximately 60 API calls per minute. Your mock fixture layer — mirroring real responses to local JSON on first fetch — must be implemented in Week 1. Without it, you will exhaust your daily allowance within the first hour of active development and lose momentum.",
    tip: "Deploy V1 to Railway by the end of Week 4 even if only company search and charts work. A live public URL with real data is immeasurably more impressive than a perfectly finished local application with no public presence. Build the URL early, then keep improving what is behind it.",
  },

  // ── PROJECT 5 ──────────────────────────────────────────────────────────────
  {
    id: "P5", color: C.p5, phase: "II", weeks: "Weeks 13–15",
    title: "Visual Multi-Agent Workspace",
    subtitle: "Event-Driven Orchestration · Role-Assigned LLMs · Human-in-the-Loop Governance",
    bottleneck: "Eliminates blocking single-thread execution stalls during long generative steps and provides structured control loops for autonomous AI chains. Single-agent architectures block on long-running model calls, cannot parallelise specialist tasks, and provide no checkpoint mechanism for human review. This project replaces linear agent execution with an event-driven orchestration matrix where each model operates as an independent worker, governed by a BullMQ job queue and supervised by explicit human approval gates.",
    overview: "The Visual Multi-Agent Workspace is an orchestration platform that coordinates three specialised local LLM agents through a shared event queue. A Decomposer receives a high-level software engineering task and breaks it into subtasks. The Architect agent (DeepSeek-r1) designs the system architecture. The Developer agent (Qwen2.5-coder) implements the code. The QA agent (Llama3.1) reviews the output and flags issues. A Next.js visual canvas displays real-time agent status (Queued / Thinking / Writing / Awaiting Review) via WebSocket, with explicit human-in-the-loop approval gates at critical checkpoints.",
    objectives: [
      "Configure a multi-session Ollama environment supporting concurrent model execution.",
      "Assign and enforce three distinct agent roles: Architect, Developer, and QA Reviewer.",
      "Implement a BullMQ job routing pipeline with sequential dependency resolution between agents.",
      "Build a Next.js visual canvas with real-time agent status tracking via Socket.io.",
      "Implement human-in-the-loop approval gates — agents pause execution and await human confirmation before proceeding.",
    ],
    stack: {
      "Language & Runtime": "Node.js 20 (TypeScript)",
      "Agent Models": "DeepSeek-r1:7b (Architect), Qwen2.5-coder:7b (Developer), Llama3.1:8b (QA)",
      "Orchestration": "BullMQ v5 (job queue with dependency chains)",
      "Real-Time UI": "Next.js 14, Socket.io, Tailwind CSS",
      "Model Interface": "Ollama REST API (streaming responses)",
      "Storage": "Redis (job state), PostgreSQL (completed task history)",
      "Infrastructure": "Docker Compose",
    },
    architecture: [
      "Task Decomposer: The orchestrator receives a high-level prompt and calls the Architect model to decompose it into a JSON array of subtasks. Each subtask carries a type (architecture/implementation/review), dependency list, and input context.",
      "Job Queue Pipeline: Subtasks are enqueued as BullMQ jobs with explicit dependency relationships. BullMQ enforces that child jobs wait for parent completion before activation. The Developer job only starts after the Architect job produces an approved architecture output.",
      "Agent Workers: Each agent role runs as an independent BullMQ worker process. Workers stream model output via the Ollama streaming API, buffering tokens and emitting partial progress events to the Socket.io server for live display.",
      "Human-in-the-Loop Gates: At defined checkpoints (e.g., after architecture design, before code execution), the pipeline emits a human-approval event to the frontend. The job enters a Waiting Approval state. The UI displays the pending output and two action buttons — Approve and Reject with Feedback. Rejection re-queues the job with the feedback appended to the model prompt.",
      "Visual Canvas: The Next.js frontend renders a swimlane diagram with one lane per agent. Each job card shows: model name, current status badge, elapsed time, and a scrollable live token stream. Approved outputs are stored in PostgreSQL as an immutable task history.",
    ],
    weekPlan: [
      { num: 13, title: "Orchestrator Foundation", tasks: ["Configure Ollama multi-model environment: DeepSeek-r1, Qwen2.5-coder, Llama3.1 all available simultaneously", "Build task decomposition endpoint: takes high-level prompt, returns JSON subtask array", "Set up BullMQ with dependency chain configuration", "Write agent worker stubs: each logs received job payload and returns a mock output", "Build Socket.io server for real-time agent event streaming"], deliverable: "Working orchestrator that decomposes a task, routes jobs through BullMQ, and emits status events" },
      { num: 14, title: "Agent Implementation + Visual Canvas", tasks: ["Implement Architect worker: streams DeepSeek-r1 response, emits token-by-token to Socket.io", "Implement Developer worker: Qwen2.5-coder with architecture output injected as system context", "Implement QA worker: Llama3.1 reviews Developer output against original requirements", "Build Next.js visual canvas: swimlane layout, live token stream cards, status badges", "Implement human-approval gate: pause pipeline, render approval UI, handle approve/reject"], deliverable: "Fully wired multi-agent pipeline with visual canvas and human approval gate operational" },
      { num: 15, title: "Polish, Testing & Documentation", tasks: ["End-to-end test: submit a real engineering task (e.g., 'build a REST CRUD API for a to-do app')", "Measure and document full pipeline execution time: decompose → architect → develop → review", "Add task history persistence to PostgreSQL: store all completed pipeline runs", "Write README: system diagram, agent roles, approval gate workflow", "Record Loom: submit a task, watch all three agents work live on the visual canvas"], deliverable: "End-to-end tested multi-agent system with history persistence, README, and Loom demo" },
    ],
    learnings: [
      "Managing continuous data flow across isolated model parameter spaces — each agent has its own context window with no shared state.",
      "Streaming buffer arrays via WebSockets: handling partial token delivery and rendering incremental output in real time.",
      "BullMQ dependency chain design: modelling producer-consumer relationships with explicit job prerequisites.",
      "Human-in-the-loop governance: designing approval gates that balance automation speed with human oversight quality.",
      "Multi-model orchestration patterns: how to structure prompts so each agent's output becomes the next agent's input without information loss.",
    ],
    metrics: [
      ["Pipeline Completion",   "E2E task in < 5 minutes",    "Orchestrator timing log",       "Week 15"],
      ["Approval Gate Accuracy","QA catches ≥ 1 issue per run","Manual review of QA outputs",  "Week 14"],
      ["UI Responsiveness",     "Token stream < 200ms lag",   "Browser DevTools network tab",  "Week 14"],
    ],
    warning: "Running three large models simultaneously (DeepSeek-r1, Qwen2.5-coder, Llama3.1) will push memory limits on most consumer machines. If RAM is constrained (< 32GB), run the pipeline sequentially rather than truly concurrently — the architectural lesson is identical.",
    tip: "This project is where your MES background becomes an advantage again. Production floor workflows — operator → engineer → quality control — follow the exact same sequential approval pattern. You already understand why each gate exists.",
  },

  // ── PROJECT 6 ──────────────────────────────────────────────────────────────
  {
    id: "P6", color: C.p6, phase: "III", weeks: "Weeks 16–18",
    title: "Transformer Architecture from Scratch",
    subtitle: "Raw Tensor Mathematics · Multi-Head Causal Attention · BPE Tokenization",
    bottleneck: "Resolves abstract comprehension caps caused by over-reliance on high-level abstraction frameworks such as Hugging Face Transformers or LangChain. Engineers who use only high-level APIs cannot debug attention score anomalies, cannot reason about gradient flow during fine-tuning, and cannot make informed architecture decisions at the component level. This project strips away all abstractions, building a decoder-only autoregressive transformer from raw PyTorch tensors.",
    overview: "This project implements a decoder-only GPT-style transformer entirely from first principles using raw PyTorch tensor operations. No Hugging Face, no LangChain, no high-level transformer libraries. Every component — scaled dot-product attention, multi-head projection, positional encoding, causal masking, layer normalisation, feedforward blocks, AdamW weight updates, and cross-entropy loss computation — is implemented manually from the mathematical definition. The model is trained on a character or BPE tokenised corpus and evaluated on next-token prediction perplexity.",
    objectives: [
      "Implement scaled dot-product multi-head causal attention from raw matrix operations.",
      "Build positional encoding matrices and understand their role in sequence awareness.",
      "Create custom character-level and BPE tokenizers without external tokenization libraries.",
      "Execute a local training script with AdamW optimisation and causal masking enforcement.",
      "Profile training dynamics: track loss curves, gradient norms, and attention weight distributions.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Deep Learning": "PyTorch 2.2 (raw tensor operations — no high-level nn.Transformer)",
      "Numerical Computing": "NumPy 1.26",
      "Tokenization": "Custom character-level tokenizer + manual BPE implementation",
      "Visualisation": "Matplotlib (loss curves, attention heat maps)",
      "Hardware": "CPU training (small model) or MPS/CUDA if available",
    },
    architecture: [
      "Tokenizer: Implements both a character-level vocabulary and a Byte-Pair Encoding (BPE) tokenizer built from scratch. The BPE implementation performs iterative merge steps on a training corpus, producing a vocabulary of configurable size. Encodes text to integer token IDs and decodes predictions back to text.",
      "Embedding Layer: Token embedding table (vocab_size × d_model) maps integer token IDs to dense vector representations. A learned positional embedding (or sinusoidal encoding, configurable) adds position-aware information. Both are implemented as nn.Embedding modules with manual initialisation.",
      "Multi-Head Causal Attention: Q, K, V projections are implemented as nn.Linear layers. The attention score matrix is computed as (Q × K^T) / sqrt(d_k). A causal mask (upper-triangular -inf matrix) is added before softmax to prevent tokens from attending to future positions. Multiple attention heads are computed in parallel and concatenated before a final projection.",
      "Feedforward Block: A two-layer MLP with GELU activation (d_model → 4×d_model → d_model) follows each attention block. Pre-layer normalisation (Pre-LN) is applied for training stability.",
      "Training Loop: AdamW optimiser with configurable learning rate, beta parameters, and weight decay. Gradient clipping (max norm 1.0) prevents exploding gradients. Loss is computed as cross-entropy between the predicted logit distribution and the shifted target token sequence (next-token prediction objective).",
    ],
    weekPlan: [
      { num: 16, title: "Tokenizer + Embeddings", tasks: ["Implement character-level tokenizer: build vocab from corpus, encode/decode functions", "Implement BPE tokenizer: iterative merge algorithm, configurable vocabulary size", "Build token embedding table and sinusoidal positional encoding matrix", "Write unit tests: encode → decode round-trip, positional encoding shape validation", "Prepare training corpus: Shakespeare or similar small text dataset"], deliverable: "Working tokenizer (char + BPE) and embedding layer with unit test coverage" },
      { num: 17, title: "Attention + Transformer Block", tasks: ["Implement scaled dot-product attention: Q/K/V projection, score matrix, causal mask, softmax", "Implement multi-head attention: parallel head computation, concatenation, output projection", "Implement feedforward block: 2-layer MLP with GELU, Pre-LN", "Assemble full decoder block: Attention + FFN + residual connections + LayerNorm", "Stack N decoder blocks into complete transformer model", "Validate: forward pass through full model, check output shape, inspect attention maps"], deliverable: "Complete decoder-only transformer model with correct forward pass and visualisable attention maps" },
      { num: 18, title: "Training Loop + Evaluation", tasks: ["Implement AdamW optimiser loop: forward pass → loss → backward → gradient clip → step", "Add learning rate warmup schedule: linear warmup over first 100 steps", "Train for 2,000 steps on corpus: plot loss curve, verify monotonic decrease", "Generate sample text at checkpoints: evaluate qualitative coherence", "Visualise attention weight heat maps for sample inputs", "Compute and report final perplexity on held-out validation set"], deliverable: "Trained transformer with loss curve plot, attention visualisations, perplexity score, and text samples" },
    ],
    learnings: [
      "Mathematical foundations of Q, K, V attention matrices: why queries search keys to retrieve weighted values.",
      "Causal masking mechanics: how the upper-triangular mask enforces the autoregressive constraint during training.",
      "Layer scaling and residual connections: why deep networks require skip connections and normalisation to train stably.",
      "Weight optimisation via Cross-Entropy and AdamW: backpropagation through attention, gradient flow dynamics.",
      "BPE tokenization algorithm: the merge iteration process, vocabulary compression ratios, and subword generalisation.",
      "Training dynamics debugging: interpreting loss curves, detecting gradient explosion, and diagnosing attention collapse.",
    ],
    metrics: [
      ["Training Loss",       "Decreasing monotonically",  "Loss curve plot (Matplotlib)",  "Week 18"],
      ["Val Perplexity",      "< 50 on held-out set",      "Perplexity computation",        "Week 18"],
      ["Attention Maps",      "Visually coherent patterns","Heat map visualisation",        "Week 17"],
      ["BPE Compression",     "≥ 30% token count reduction","Corpus encode stats",           "Week 16"],
    ],
    warning: "Do not use nn.MultiheadAttention or nn.TransformerDecoderLayer — these are wrapper modules that hide the mathematical operations you need to understand. Implement attention as raw tensor operations: torch.matmul for Q×K^T, torch.softmax with the causal mask applied as additive -inf, manual reshape and permute for multi-head splitting.",
    tip: "Train on a small, well-understood corpus — Shakespeare's complete works (~1MB) is the classic choice. This keeps training time manageable on CPU while producing text samples interesting enough to qualitatively evaluate. Resist the temptation to use a GPU-scale dataset; the goal is understanding, not capability.",
  },

  // ── PROJECT 7 ──────────────────────────────────────────────────────────────
  {
    id: "P7", color: C.p7, phase: "III", weeks: "Weeks 19–21",
    title: "Vector Indexing Engine",
    subtitle: "Spatial Indexing Algorithms · IVF Clustering · HNSW Graph Routing",
    bottleneck: "Addresses the fundamental algorithmic failure of B-Tree indexing under high-dimensional vector data. Traditional database indices optimise for exact equality and range queries over scalar dimensions. At 768+ vector dimensions, B-Tree traversal degenerates to linear scan complexity. This project implements two dominant approximate nearest-neighbour (ANN) algorithms — Inverted File Index (IVF) with k-means routing and Hierarchical Navigable Small World (HNSW) graph search — from raw NumPy operations.",
    overview: "This project builds a standalone vector indexing library implementing the two most widely deployed ANN algorithms in production vector databases. IVF partitions the vector space into k Voronoi cells using k-means clustering, then at query time probes the nprobe nearest centroids to find candidate vectors. HNSW builds a hierarchical proximity graph where nodes connect to their approximate nearest neighbours at multiple granularity levels, enabling greedy beam search with O(log N) average complexity. The project benchmarks both algorithms against brute-force linear scan across datasets of increasing dimensionality.",
    objectives: [
      "Implement a brute-force linear scan baseline for exact nearest-neighbour search.",
      "Build an IVF index using manual k-means clustering with Lloyd's algorithm.",
      "Implement HNSW graph construction and greedy beam search without external libraries.",
      "Benchmark IVF vs. HNSW vs. brute-force: queries-per-second, recall rate, and build time.",
      "Visualise the Voronoi cell structure (IVF) and graph layer structure (HNSW) for a 2D toy dataset.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Numerical Computing": "NumPy 1.26 (all vector arithmetic, no external ANN libraries)",
      "Visualisation": "Matplotlib (Voronoi diagrams, graph layer plots, QPS vs. recall curves)",
      "Benchmarking": "Python time module, custom precision/recall computation",
      "Reference Dataset": "GloVe-25d word vectors (small, fast) or random synthetic high-dim data",
    },
    architecture: [
      "Brute-Force Baseline: Computes cosine distance from the query vector to every stored vector. Implemented as a vectorised NumPy matrix operation (query @ corpus.T) for efficiency. Returns exact top-K results. Used as ground truth for recall measurement.",
      "IVF Index — Build Phase: Runs Lloyd's k-means algorithm for a configurable number of iterations and cluster count k. Assigns each corpus vector to its nearest centroid. Stores an inverted list: a mapping from centroid ID to the list of vector IDs in that Voronoi cell.",
      "IVF Index — Query Phase: Computes distance from the query vector to all k centroids. Selects the nprobe nearest centroids. Retrieves all vectors from those nprobe inverted lists. Performs an exact search over the candidate set and returns the top-K results.",
      "HNSW Index — Build Phase: Constructs a multi-layer proximity graph. Each new vector is inserted at a randomly assigned maximum layer. Greedy search from the top layer finds the entry point's nearest neighbours at each layer. Bidirectional edges are added between the new vector and its M nearest neighbours at each layer.",
      "HNSW Index — Query Phase: Entry point is the globally designated top-layer node. Greedy descent: at each layer, explore neighbours and move to any that are closer to the query. At the bottom layer (layer 0), a beam search with configurable ef parameter expands the candidate set before returning top-K results.",
    ],
    weekPlan: [
      { num: 19, title: "Brute-Force + IVF Build", tasks: ["Implement brute-force cosine similarity search as vectorised NumPy baseline", "Implement Lloyd's k-means algorithm: random init, assignment step, centroid update, convergence check", "Build IVF index: assign corpus vectors to nearest centroids, construct inverted lists", "Build IVF query: centroid distance ranking, nprobe selection, candidate set retrieval", "Benchmark IVF vs. brute-force: QPS and recall@10 on 100K vector dataset"], deliverable: "IVF index with benchmarked QPS and recall vs. brute-force baseline" },
      { num: 20, title: "HNSW Implementation", tasks: ["Implement HNSW node and layer data structures using Python dicts", "Build HNSW insert: layer assignment, greedy entry search, M-nearest-neighbour edge creation", "Build HNSW query: top-layer entry, layer-by-layer greedy descent, bottom-layer beam search with ef", "Validate: HNSW recall@10 ≥ 95% on synthetic dataset", "Benchmark HNSW vs. IVF vs. brute-force: QPS, recall, build time"], deliverable: "HNSW index with recall ≥ 95% and QPS benchmark comparison table against IVF and brute-force" },
      { num: 21, title: "Visualisation + Analysis", tasks: ["Visualise Voronoi cell structure of IVF using 2D PCA-projected data (Matplotlib)", "Visualise HNSW layer structure: draw graph edges at each layer for a 50-node toy dataset", "Plot QPS vs. recall tradeoff curve for both IVF (varying nprobe) and HNSW (varying ef)", "Write analytical report: when to choose IVF vs. HNSW vs. pgvector for a given workload", "Package as importable Python library with clean API: IndexIVF, IndexHNSW classes"], deliverable: "Complete visualisation suite, tradeoff analysis report, and importable Python indexing library" },
    ],
    learnings: [
      "High-dimensional proximity scanning: why Euclidean distance loses discriminative power in high dimensions (curse of dimensionality).",
      "Graph routing loops: HNSW's hierarchical navigation exploits small-world graph properties for logarithmic average search complexity.",
      "Voronoi decomposition and k-means centroid assignment as a space partitioning strategy for approximate search.",
      "The precision-speed tradeoff in ANN algorithms: nprobe in IVF and ef in HNSW as configurable accuracy knobs.",
      "Build-time vs. query-time complexity: IVF has faster queries for large nprobe but slower recall; HNSW has high build cost but excellent query performance.",
    ],
    metrics: [
      ["HNSW Recall@10",    "≥ 95%",              "vs. brute-force ground truth", "Week 20"],
      ["IVF QPS",           "≥ 10× brute-force",  "Python time benchmarking",     "Week 19"],
      ["HNSW QPS",          "≥ 20× brute-force",  "Python time benchmarking",     "Week 20"],
      ["Build Time (HNSW)", "< 60s for 100K vecs","Wall-clock time measurement",  "Week 20"],
    ],
    warning: "Do not use faiss, annoy, or hnswlib — these are the libraries you are reimplementing. All vector arithmetic must be raw NumPy operations. The value of this project is in understanding the algorithm, not the output. Using a library defeats the entire purpose.",
    tip: "Test your IVF implementation with k=10 centroids on a 1,000-vector dataset before scaling to 100K. Visualise the Voronoi cells using 2D synthetic data — seeing the spatial partitioning makes the algorithm intuitive. Debug spatially before benchmarking numerically.",
  },

  // ── PROJECT 8 ──────────────────────────────────────────────────────────────
  {
    id: "P8", color: C.p8, phase: "III", weeks: "Weeks 22–25",
    title: "Quantization & Model Inference Engine",
    subtitle: "FP32 to INT4 Weight Compression · KV Cache Architecture · VRAM Optimisation",
    bottleneck: "Resolves VRAM capacity exhaustion, memory bus saturation, and redundant compute during deep sequence loops. Modern LLMs store weights as 32-bit floating-point values, consuming approximately 4 bytes per parameter. A 7B parameter model at FP32 requires 28GB of VRAM — far exceeding consumer GPU capacity. Quantization reduces each weight's bit-width representation, trading a small, measurable accuracy loss for dramatic reductions in memory footprint and memory bandwidth consumption.",
    overview: "This project implements a custom quantization engine that converts model weights from 32-bit floating-point (FP32) to 8-bit integer (INT8) and 4-bit integer (INT4) representations. Alongside quantization, the project builds a Key-Value (KV) Cache management system that stores computed attention keys and values across generation steps, eliminating redundant forward-pass computation for the prefix context. A telemetry dashboard tracks compression ratios, memory footprint, and generation throughput across precision levels.",
    objectives: [
      "Implement symmetric and asymmetric quantization schemes for FP32 to INT8 and INT4 weight compression.",
      "Build a per-channel quantization calibration pipeline using representative data samples.",
      "Design and implement a KV Cache data structure for attention key-value pair storage.",
      "Build a low-level Python/C++ inference layer that executes quantized weight matrix multiplications.",
      "Monitor and chart compression ratios, memory usage, and semantic accuracy across precision levels.",
    ],
    stack: {
      "Language": "Python 3.11, C++ (via ctypes or pybind11 for INT4 kernel)",
      "Deep Learning": "PyTorch 2.2 (base model from P6), NumPy",
      "Systems Tools": "tracemalloc (memory profiling), Python ctypes",
      "Visualisation": "Matplotlib (accuracy vs. compression curve)",
      "Reference Model": "Project 6 transformer or a small Hugging Face GPT-2 (124M) for baseline",
    },
    architecture: [
      "Symmetric Quantization: Maps weight tensors from the FP32 range [-max, max] to INT8 [-127, 127] using a single scale factor: scale = max(|W|) / 127. Dequantisation reverses the mapping. Symmetric quantization is simpler but less accurate than asymmetric for skewed weight distributions.",
      "Asymmetric Quantization (INT4): Uses separate scale and zero-point values per quantization group: Q = round((W / scale) + zero_point). Group-wise quantization (128 weights per group) minimises quantization error by adapting the mapping to local weight statistics.",
      "Per-Channel Calibration: Runs a forward pass over a calibration dataset (small representative sample). Records the activation range of each linear layer's output. Uses these observed ranges to set per-channel scale factors rather than relying solely on weight statistics.",
      "KV Cache: During autoregressive generation, the attention mechanism computes Key and Value projections for every token in the prefix context at every generation step. The KV Cache stores these projections in a pre-allocated tensor buffer indexed by [batch, layer, head, sequence_position]. At each new generation step, only the new token's K and V projections are computed; all prior entries are read from cache.",
      "INT4 Kernel (C++): The memory bandwidth bottleneck in LLM inference is loading weights from DRAM. An INT4 kernel packs two weights per byte, halving memory bandwidth consumption. The C++ kernel dequantises weights on-the-fly during the matrix multiply, executing the operation at FP32 arithmetic while transferring data at INT4 bandwidth.",
    ],
    weekPlan: [
      { num: 22, title: "FP32 → INT8 Quantization", tasks: ["Implement symmetric INT8 quantization: scale computation, quantize, dequantize functions", "Apply to all linear layer weight tensors in the P6 transformer model", "Measure: model size before vs. after (MB), memory usage via tracemalloc", "Evaluate: run 20 benchmark questions from P1 — compare INT8 vs. FP32 output quality", "Visualise weight distribution before and after quantization with Matplotlib histogram"], deliverable: "INT8 quantized model with memory footprint measurement and accuracy comparison table" },
      { num: 23, title: "INT4 + Per-Channel Calibration", tasks: ["Implement asymmetric INT4 quantization with group-wise quantization (group size 128)", "Build per-channel calibration pipeline: forward pass over calibration data, record activation ranges", "Apply calibrated INT4 quantization to all weight tensors", "Compare: FP32 vs. INT8 vs. INT4 — model size MB, perplexity, generation quality", "Write simple C++ INT4 dequantization function callable via Python ctypes"], deliverable: "INT4 quantized model with C++ dequantization kernel and three-way accuracy comparison" },
      { num: 24, title: "KV Cache Implementation", tasks: ["Design KV Cache buffer: pre-allocated tensor [batch, n_layers, n_heads, max_seq, head_dim]", "Modify transformer forward pass to write K,V projections to cache at each generation step", "At each step, concatenate cached K,V with new token's K,V before attention computation", "Measure tokens/sec: generation with KV Cache vs. without KV Cache (prefix context of 512 tokens)", "Validate: KV Cache and no-cache produce identical output text for same input"], deliverable: "KV Cache implementation with measured generation throughput improvement across prefix lengths" },
      { num: 25, title: "Telemetry Dashboard + Analysis", tasks: ["Build Matplotlib dashboard: 4-panel chart (model size, memory, perplexity, tokens/sec vs. precision)", "Plot compression ratio vs. semantic accuracy degradation curve", "Document: at what INT4 configuration does quality degradation become unacceptable?", "Write analysis report: compute-bound vs. memory-bound inference and when quantization helps", "Package quantization utilities as reusable Python module for use in P12"], deliverable: "Four-panel telemetry dashboard, analysis report, and reusable quantization module" },
    ],
    learnings: [
      "Range scale values and rounding drift: how quantization error accumulates through layers and compounds across depth.",
      "The shift from compute-bound to memory-bandwidth-bound inference: why LLMs on consumer GPUs are limited by VRAM bandwidth, not FLOP capacity.",
      "KV Cache mechanics: why recomputing prefix attention at every generation step is redundant and how pre-allocation solves it.",
      "Group-wise vs. per-tensor quantization: the accuracy-compression tradeoff and when each granularity is appropriate.",
      "INT4 packing arithmetic: storing two 4-bit values per byte and the bit manipulation required to unpack them efficiently.",
    ],
    metrics: [
      ["INT4 Model Size",       "~75% reduction vs. FP32",    "Model parameter byte count",    "Week 23"],
      ["KV Cache Speedup",      "≥ 3× tokens/sec at len 512", "Generation timing benchmark",   "Week 24"],
      ["INT4 Quality Delta",    "< 5% perplexity increase",   "Perplexity on validation set",  "Week 23"],
      ["Memory Footprint",      "INT4 ≤ 25% of FP32 usage",   "tracemalloc measurement",       "Week 23"],
    ],
    warning: "INT4 quantization of a poorly trained model will produce incoherent output — the model must be reasonably trained before quantization degrades it further. If the P6 transformer quality is too low, use a Hugging Face GPT-2 (124M) as the target model for quantization experiments. This is explicitly permitted for this project.",
    tip: "The KV Cache implementation is the most practically impactful component in this project. In production inference servers like vLLM and TensorRT-LLM, KV Cache management is the primary bottleneck for batch throughput. Understanding it from scratch gives you rare depth for a self-taught engineer.",
  },

  // ── PROJECT 9 ──────────────────────────────────────────────────────────────
  {
    id: "P9", color: C.p9, phase: "IV", weeks: "Weeks 26–27",
    title: "Sparse Mixture of Experts (MoE) Transformer",
    subtitle: "Expert Routing Gates · Auxiliary Load Balancing · Sparse Activation Patterns",
    bottleneck: "Eliminates dense computing inefficiencies where every model parameter activates for every input token. In a standard dense transformer, all feedforward parameters process every token regardless of relevance, wasting compute on non-specialist operations. Mixture of Experts replaces a single large feedforward layer with N parallel expert networks and a learned gating function that routes each token to only the top-K most relevant experts, dramatically reducing activated parameters per forward pass while maintaining total model capacity.",
    overview: "This project extends the Project 6 transformer by replacing the feedforward block with a Sparse MoE layer. A learnable gating network scores each token against all available experts and selects the top-K (typically K=2) for activation. An Auxiliary Load Balancing Loss term penalises expert utilisation imbalance, preventing the model from routing all tokens to a single expert and making the other experts idle. The implementation builds expert networks as independent linear modules and validates sparse activation patterns with visualisation tools.",
    objectives: [
      "Build a routing gating network with Softmax probabilities over N expert modules.",
      "Implement top-K expert selection with configurable K and expert count N.",
      "Add Auxiliary Load Balancing Loss to enforce uniform expert utilisation during training.",
      "Integrate the MoE layer into the P6 transformer as a drop-in feedforward replacement.",
      "Visualise expert utilisation distributions across training batches.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Framework": "PyTorch 2.2 (extending P6 transformer codebase)",
      "Visualisation": "Matplotlib (expert utilisation bar charts, routing entropy plots)",
      "Base": "Project 6 transformer architecture",
    },
    architecture: [
      "Gating Network: A linear layer (d_model → N) followed by Softmax produces a probability distribution over N experts for each input token. Top-K selection retains only the K highest-probability experts per token; all others are zeroed.",
      "Expert Networks: N independent feedforward modules, each with the same architecture as the P6 FFN block (d_model → 4×d_model → d_model with GELU). Only the K selected experts receive a forward pass per token.",
      "Auxiliary Load Balancing Loss: Computed as the dot product of the average expert selection probability and the average fraction of tokens routed to each expert. Added to the primary cross-entropy loss with a configurable coefficient α (typically 0.01). This penalises routing collapse — the degenerate state where all tokens route to one expert.",
      "Token Capacity: Each expert has a configurable maximum capacity (tokens per batch). Tokens routed to an overloaded expert beyond capacity are dropped (with position-aware overflow handling). This enforces parallelism in distributed settings.",
    ],
    weekPlan: [
      { num: 26, title: "MoE Layer Implementation", tasks: ["Build N expert FFN modules as independent nn.Module instances", "Implement gating network: linear projection + Softmax over N experts", "Implement top-K selection: retain K expert probabilities, zero others, renormalise", "Compute weighted sum of K expert outputs per token", "Implement Auxiliary Load Balancing Loss with configurable α coefficient", "Unit test: verify exactly K experts activate per token, others produce zero output"], deliverable: "MoE layer with top-K gating, N expert networks, and Auxiliary Loss — unit tested" },
      { num: 27, title: "Integration + Utilisation Analysis", tasks: ["Replace FFN block in P6 transformer with MoE layer (drop-in swap)", "Train for 500 steps on corpus with α=0.01: verify loss decreases, aux loss converges", "Compare training dynamics: MoE loss curve vs. dense FFN loss curve from P6", "Plot expert utilisation bar chart across 1,000 training batches: verify near-uniform distribution", "Ablation: train without Auxiliary Loss — observe routing collapse to 1-2 experts", "Document: activated-parameter count MoE vs. dense at equivalent total capacity"], deliverable: "MoE transformer with uniform utilisation (visualised), routing collapse ablation, and capacity comparison" },
    ],
    learnings: [
      "Sparse network computation logic: how conditional activation reduces FLOPs per forward pass while maintaining total model parameter count.",
      "Auxiliary Loss formula derivation: why the dot product of selection frequency and routing probability enforces balance.",
      "Softmax gating probability interpretation: the relationship between gate temperature and routing sharpness.",
      "Expert capacity constraints in distributed systems: why token dropping is preferred over dynamic batching overhead.",
      "Routing collapse failure mode: what causes all tokens to route to one expert and how Auxiliary Loss prevents it.",
    ],
    metrics: [
      ["Expert Utilisation",    "± 10% uniform distribution",  "Bar chart across 1K batches",   "Week 27"],
      ["Aux Loss Convergence",  "Stabilises within 100 steps", "Loss curve plot",               "Week 27"],
      ["Activated Params/Token","K/N × total FFN params",      "Manual calculation + assertion", "Week 26"],
    ],
    warning: "Routing collapse (all tokens routing to one expert) is the dominant failure mode. If the Auxiliary Loss coefficient α is set too low, the model learns to ignore most experts within the first 50 training steps. Start with α=0.01; increase to 0.1 if collapse is observed. Plot expert utilisation every 100 steps during training.",
    tip: "This is conceptually the hardest project in Phase IV. Ground yourself in a simple case first: N=4 experts, K=1 (top-1 routing), trained for 200 steps. Once you verify that the Auxiliary Loss works as expected at small scale, scale to N=8 and K=2.",
  },

  // ── PROJECT 10 ─────────────────────────────────────────────────────────────
  {
    id: "P10", color: C.p10, phase: "IV", weeks: "Weeks 28–29",
    title: "Speculative Decoding Engine",
    subtitle: "Draft-Verify Parallel Inference · Token Acceptance Loops · Throughput Optimisation",
    bottleneck: "Addresses memory-bandwidth constraints that force hardware stalls while the model reads parameter weights step-by-step during sequential generation. In standard autoregressive decoding, the large target model generates exactly one token per forward pass, requiring a full weight reload for each step. Speculative decoding decouples token proposal from token verification, enabling the small draft model to batch-speculate multiple tokens in parallel while the target model verifies them in a single forward pass.",
    overview: "This project implements a speculative decoding system pairing a small 1B-parameter draft model with a 7B-parameter target model. The draft model generates a speculative lookahead of γ tokens per cycle. The target model evaluates all γ draft tokens in a single parallel forward pass. Tokens are accepted probabilistically using a residual distribution sampling rule derived from the ratio of target and draft probabilities. Accepted tokens are output; the first rejected token triggers the target model to sample a corrected replacement, and the cycle restarts. The system benchmarks tokens-per-second improvement over sequential target-model-only generation.",
    objectives: [
      "Build a lookahead generation pipeline using a small draft model (Ollama 1B).",
      "Implement statistical residual acceptance sampling: P_accept = min(1, P_target / P_draft).",
      "Benchmark tokens-per-second improvement across varying lookahead depth γ values.",
      "Handle edge cases: complete acceptance, partial acceptance, and full rejection cycles.",
      "Build a dashboard comparing baseline sequential generation vs. speculative decoding throughput.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Model Interface": "Ollama REST API (streaming + logprobs endpoint) or local PyTorch models",
      "Numerical Computing": "NumPy (probability computation), Python math module",
      "Visualisation": "Matplotlib (tokens/sec vs. γ chart, acceptance rate per position)",
      "Benchmarking": "Python time module, custom throughput tracker",
    },
    architecture: [
      "Draft Phase: The draft model generates γ candidate tokens autoregressively. At each draft step, the model's top-1 logit prediction is recorded along with the full probability distribution over the vocabulary. The γ draft tokens and their corresponding draft probability distributions are stored as a speculative batch.",
      "Verify Phase: The target model receives the full input context plus all γ draft tokens concatenated. In a single forward pass, it produces γ+1 output logit distributions — one for each draft token position plus one for the next token after accepting all γ tokens.",
      "Acceptance Loop: For each draft position i (0 to γ-1), compute acceptance probability as min(1, P_target(token_i) / P_draft(token_i)). Sample a Uniform(0,1) random variable u. If u < P_accept, token_i is accepted and the loop advances. If u ≥ P_accept, token_i is rejected; the target model samples a corrected token from the residual distribution max(0, P_target - P_draft) / Z and the cycle resets.",
      "Throughput Analysis: Tokens-per-second is measured for three configurations: target model only (baseline), speculative with γ=4, and speculative with γ=8. The optimal γ balances the draft model's speed advantage against its acceptance rate — high rejection rates with large γ can negate the throughput gains.",
    ],
    weekPlan: [
      { num: 28, title: "Draft Pipeline + Acceptance Logic", tasks: ["Interface with Ollama to extract per-token logit distributions (logprobs endpoint)", "Implement draft phase: generate γ tokens with draft model, store token IDs and probability distributions", "Implement target verify phase: single forward pass evaluating all γ draft positions", "Implement acceptance loop: P_accept computation, uniform sampling, rejection correction", "Test with γ=4 on 10 example prompts: log acceptance rate per position"], deliverable: "Working speculative decoding pipeline with logged per-position acceptance rates" },
      { num: 29, title: "Benchmarking + Dashboard", tasks: ["Benchmark tokens/sec: baseline target-only generation on 20 prompts (200 tokens each)", "Benchmark speculative decoding: γ=2, 4, 8 on identical prompts", "Compute and plot: acceptance rate vs. position in lookahead (expect decay at high γ)", "Plot tokens/sec vs. γ curve — identify optimal γ for this draft/target pair", "Validate output quality: compare generated text between speculative and sequential for identical seeds", "Write analysis: at what rejection rate does speculative decoding become slower than baseline?"], deliverable: "Benchmark dashboard (tokens/sec vs. γ), acceptance rate analysis, and breakeven rejection rate report" },
    ],
    learnings: [
      "Converting sequential token generation to parallel verification batches — why the target model's throughput is bandwidth-bound, not compute-bound.",
      "Residual distribution sampling: how the rejection correction maintains the exact same output distribution as the target model alone.",
      "Context truncation in acceptance loops: managing position offsets when cycles restart at a rejection point.",
      "The γ optimisation problem: why more lookahead is not always better and how acceptance rate governs the performance curve.",
      "Memory bandwidth as the primary inference bottleneck: why larger models benefit more from speculative decoding than smaller ones.",
    ],
    metrics: [
      ["Throughput Improvement","≥ 2× tokens/sec over baseline","Python time benchmark",        "Week 29"],
      ["Output Equivalence",    "Identical distribution",       "Statistical equivalence test", "Week 29"],
      ["Optimal γ",             "Empirically determined",       "γ vs. tokens/sec curve",       "Week 29"],
      ["Acceptance Rate",       "≥ 70% at γ=4",                "Per-position logging",         "Week 28"],
    ],
    warning: "Logprob extraction from Ollama's API requires the logprobs option in the generate endpoint. Verify this is available in your Ollama version before starting Week 28. If logprobs are unavailable, this project can be implemented using the P6 PyTorch model directly — the acceptance math is identical regardless of the model source.",
    tip: "The key insight of speculative decoding is that the output distribution is mathematically identical to running the target model alone. This is not an approximation — it is provably exact under the acceptance sampling rule. Understanding why requires reading the residual distribution derivation carefully before implementing it.",
  },

  // ── PROJECT 11 ─────────────────────────────────────────────────────────────
  {
    id: "P11", color: C.p11, phase: "IV", weeks: "Weeks 30–31",
    title: "Automated Evaluation (LLM-as-a-Judge) Framework",
    subtitle: "Pass@K Metrics · Sandboxed Code Execution · Chain-of-Thought Grading",
    bottleneck: "Replaces subjective, unquantifiable qualitative assessment with statistical validation metrics across prompt modifications. Evaluating LLM output quality by reading responses is unscalable, inconsistent across evaluators, and blind to regression when prompts or models change. This project implements a closed-loop automated evaluation pipeline where a judge LLM scores outputs using structured chain-of-thought rubrics, and generated code is tested by executing it in isolated Docker containers with assertion-based unit tests.",
    overview: "The Automated Evaluation Framework treats LLM output assessment as a software quality assurance problem. For coding tasks, it generates multiple candidate solutions using the model under test, executes each in a sandboxed Docker container against a hidden unit test suite, and computes Pass@K probabilities. For open-ended text tasks, a separate judge LLM evaluates outputs against a structured rubric covering accuracy, coherence, instruction following, and reasoning depth. The framework produces regression dashboards enabling prompt engineers to objectively track quality improvements or degradations across model and prompt versions.",
    objectives: [
      "Build a Pass@K calculator for code generation tasks: sample K solutions, report P(at least 1 passes).",
      "Configure Docker SDK-based sandboxed containers for safe, isolated code execution.",
      "Implement an LLM-as-a-judge pipeline with structured chain-of-thought scoring rubrics.",
      "Build a regression dashboard tracking evaluation scores across prompt and model versions.",
      "Design and evaluate a benchmark suite of 30 diverse test cases across coding and reasoning domains.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Code Execution": "Docker SDK for Python (isolated container execution)",
      "Judge Model": "Local Ollama model (Qwen2.5-coder or DeepSeek-r1)",
      "Visualisation": "Matplotlib (Pass@K curves, score distribution histograms)",
      "Benchmarking": "Custom test suite (30 diverse evaluation prompts)",
      "Storage": "PostgreSQL (evaluation run history, score time series)",
    },
    architecture: [
      "Pass@K Computation: For each coding task in the benchmark suite, K candidate solutions are sampled from the model under test (with temperature > 0 for diversity). Each candidate is executed in an isolated Docker container (Python 3.11 image) against a hidden unit test file. The unbiased Pass@K estimator is: Pass@K = 1 - C(n-c, K) / C(n, K), where n is total samples and c is correct samples.",
      "Sandboxed Execution: The Docker SDK creates a fresh container per code candidate, mounts a temp directory containing the generated code and test file, executes python -m pytest with a timeout, and captures stdout, stderr, exit code, and test result summary. The container is torn down immediately after execution. No network access or filesystem access outside the mounted temp directory is permitted.",
      "LLM-as-a-Judge Pipeline: For non-code tasks, a judge model receives the original prompt, the model's response, and a structured evaluation rubric. The rubric specifies 4–6 criteria (accuracy, completeness, instruction following, reasoning depth, conciseness) each scored 1–5. The judge model produces its assessment in chain-of-thought format, then outputs a final JSON score object.",
      "Regression Dashboard: Each evaluation run is stored in PostgreSQL with model name, prompt version, task ID, score, and timestamp. The dashboard plots score time series per task category, identifies regression tasks (score drop ≥ 10% from baseline), and highlights improvement tasks for prompt change attribution.",
    ],
    weekPlan: [
      { num: 30, title: "Pass@K + Docker Sandbox", tasks: ["Design 30-task benchmark suite: 15 coding tasks (with hidden unit tests), 15 reasoning tasks", "Configure Docker SDK: create container, mount temp directory, execute with timeout, capture results", "Implement Pass@K estimator using the unbiased estimator formula", "Sample K=5 solutions per coding task, execute all 75 candidates, log pass/fail", "Plot Pass@1 vs. Pass@5 per task category: identify model weaknesses by domain"], deliverable: "Pass@K evaluation pipeline running against 15 coding tasks with sandboxed Docker execution" },
      { num: 31, title: "LLM-as-a-Judge + Regression Dashboard", tasks: ["Write structured evaluation rubrics for 15 reasoning tasks (4-criteria JSON schema)", "Implement judge pipeline: prompt construction with rubric, model call, response parsing, score extraction", "Validate judge consistency: run same task 3 times, measure score variance (expect < 0.5 std)", "Build PostgreSQL schema for evaluation run history", "Build regression dashboard: score time series per task, regression/improvement detection", "Run full 30-task benchmark suite: document overall Pass@K and judge score distribution"], deliverable: "Complete evaluation framework: Pass@K pipeline, LLM-as-judge, regression dashboard, 30-task benchmark results" },
    ],
    learnings: [
      "Formulating objective statistical evaluation metrics: why Pass@K is preferred over Pass@1 for estimating model capability.",
      "Docker SDK-based sandboxing: isolating untrusted code execution while maintaining observability through captured output streams.",
      "Structured chain-of-thought grading: how rubric design affects judge consistency and score calibration.",
      "Regression testing for AI systems: treating prompt and model changes as software versions requiring quantitative quality gates.",
      "Judge model calibration: measuring inter-run variance to establish confidence intervals around evaluation scores.",
    ],
    metrics: [
      ["Judge Consistency",     "Score std < 0.5 across 3 runs","3× repeated run comparison",  "Week 31"],
      ["Container Startup",     "< 3s per execution",           "Docker SDK timing",            "Week 30"],
      ["Benchmark Coverage",    "30 diverse tasks complete",    "Task suite execution log",     "Week 31"],
      ["Pass@5 vs. Pass@1 Gap", "Measurable delta observed",    "Pass@K calculator output",     "Week 30"],
    ],
    warning: "Docker container cold starts can take 15–30 seconds the first time an image is pulled. Pre-pull the python:3.11-slim image before beginning load testing to avoid distorting execution time measurements. Add container_timeout=30 for all code execution calls to prevent infinite-loop candidates from hanging the evaluation pipeline.",
    tip: "The regression dashboard is the most practically deployable artifact from this project. After completing P12 (fine-tuning), use this framework to objectively measure whether fine-tuning improved or degraded performance on your benchmark suite. The two projects were designed to work together.",
  },

  // ── PROJECT 12 ─────────────────────────────────────────────────────────────
  {
    id: "P12", color: C.p12, phase: "IV", weeks: "Weeks 32–34",
    title: "Knowledge Distillation & Fine-Tuning Pipeline",
    subtitle: "LoRA Low-Rank Adapters · PEFT Framework · Adapter Merging · 6GB VRAM Constraint",
    bottleneck: "Overcomes the VRAM overhead boundary required to fully retrain foundation model parameters. Full fine-tuning of a 7B parameter model requires approximately 112GB of VRAM for model weights, gradients, and optimiser states — well beyond any consumer GPU. Parameter-Efficient Fine-Tuning (PEFT) freezes all base model parameters and trains only small, low-rank decomposition matrices injected at each linear layer. LoRA reduces trainable parameters by 99% while retaining 95%+ of full fine-tuning performance on domain-specific tasks.",
    overview: "This project implements a complete LoRA-based fine-tuning pipeline for a local language model. A synthetic training dataset is collected using a high-capability teacher model (GPT-4o or Qwen-Max API). A LoRA wrapper is applied to all query and value projection matrices, freezing base weights and introducing low-rank A×B matrices with configurable rank r (typically 4–16) and alpha scaling. The pipeline trains entirely within a 6GB VRAM budget, then merges fine-tuned adapters back into the base model weights for efficient inference. P11's evaluation framework is used to objectively measure fine-tuning effectiveness.",
    objectives: [
      "Collect 500–1,000 synthetic training examples from a teacher model via API.",
      "Implement a LoRA wrapper applying rank-r low-rank decomposition to target linear layers.",
      "Train the LoRA adapter using Hugging Face PEFT within a 6GB VRAM constraint.",
      "Merge fine-tuned adapters back into base model weights using the A×B matrix sum.",
      "Evaluate fine-tuned model using the P11 evaluation framework — report Pass@K and judge score improvement.",
    ],
    stack: {
      "Language": "Python 3.11",
      "Fine-Tuning": "Hugging Face PEFT (LoraConfig, get_peft_model), Transformers",
      "Training": "PyTorch 2.2, Hugging Face TRL (SFTTrainer)",
      "Dataset": "Custom synthetic dataset from teacher model (JSON format)",
      "Evaluation": "P11 automated evaluation framework (Pass@K + LLM-as-judge)",
      "Hardware Target": "≤ 6GB VRAM (or CPU with gradient checkpointing if no GPU)",
    },
    architecture: [
      "Dataset Collection: A structured prompting template is used to query a teacher model (high-capability API model). Each example consists of an instruction and a high-quality reference response. Examples are deduplicated, filtered for length, and split into training (80%) and validation (20%) sets.",
      "LoRA Architecture: For each target linear layer W (shape d_out × d_in), two low-rank matrices are introduced: A (shape r × d_in) initialised with random Gaussian weights, and B (shape d_out × r) initialised to zero. The adapted output is computed as W×x + (B×A)×x × (alpha/r). At initialisation, the LoRA contribution is exactly zero (because B=0), preserving the base model's behaviour.",
      "Training Loop: Hugging Face SFTTrainer handles data collation, gradient accumulation, and mixed-precision training. Gradient checkpointing is enabled to reduce activation memory overhead. Only the A and B LoRA matrices receive gradient updates; all base model parameters have requires_grad=False. AdamW with paged offload (bitsandbytes) keeps optimiser states in CPU RAM.",
      "Adapter Merging: After training, each LoRA pair (A, B) is merged into the corresponding base weight W as: W_merged = W + B×A×(alpha/r). The merged model has the same parameter count and inference cost as the original base model with zero LoRA overhead at inference time.",
      "Evaluation: The merged model is evaluated against the P11 benchmark suite using Pass@K for coding tasks and LLM-as-judge for reasoning tasks. Results are compared against the untuned baseline, demonstrating measurable improvement on the target task domain.",
    ],
    weekPlan: [
      { num: 32, title: "Dataset Collection + LoRA Setup", tasks: ["Design 20 domain-specific task templates (coding, reasoning, or domain Q&A)", "Query teacher model API for 500+ reference examples: instruction + high-quality response pairs", "Filter dataset: remove duplicates, length outliers, and incoherent responses", "Apply Hugging Face LoraConfig to target model: r=8, alpha=16, target_modules=['q_proj','v_proj']", "Verify trainable parameter count: should be < 1% of total parameters", "Test forward pass with LoRA: verify output shape, confirm base weights frozen"], deliverable: "Filtered training dataset (500+ examples) and configured LoRA model with verified frozen base weights" },
      { num: 33, title: "Training + Adapter Merging", tasks: ["Configure SFTTrainer: batch_size=4, gradient_accumulation=8, gradient_checkpointing=True", "Run training for 3 epochs: monitor training loss, validation loss, and VRAM usage", "Plot training curves: verify convergence, check for overfitting (val loss diverging from train)", "Generate sample outputs from LoRA-adapted model at checkpoints: assess qualitative improvement", "Merge LoRA adapters: W_merged = W + B×A×(alpha/r) for all target layers", "Verify merged model parameter count matches original base model exactly"], deliverable: "Trained LoRA adapter with converged loss curves and merged model ready for evaluation" },
      { num: 34, title: "Evaluation + Documentation", tasks: ["Run P11 evaluation suite on merged fine-tuned model: Pass@K for coding tasks", "Run P11 LLM-as-judge on reasoning tasks: compare scores to untuned baseline", "Compute improvement delta: percentage gain in Pass@K and judge scores post fine-tuning", "Ablation: test r=4 vs. r=8 vs. r=16 — plot accuracy vs. trainable parameter count tradeoff", "Write technical report: dataset collection methodology, LoRA config rationale, evaluation results", "Package pipeline as reusable Hugging Face-compatible fine-tuning script"], deliverable: "Evaluation report with measurable improvement over baseline, ablation results, and reusable training script" },
    ],
    learnings: [
      "Low-rank matrix functions: why B×A with small r approximates the full weight update at a fraction of the parameter cost.",
      "Backpropagation through frozen layers: how gradient flow stops at frozen base weights and only updates A and B matrices.",
      "Knowledge distillation theory: why teacher-generated synthetic data produces better fine-tuning outcomes than raw web data for narrow tasks.",
      "VRAM budget management: gradient checkpointing, paged AdamW, and mixed-precision training as orthogonal memory reduction strategies.",
      "Adapter merging arithmetic: why the merged model has zero inference overhead and how the alpha/r scaling factor normalises the contribution.",
    ],
    metrics: [
      ["Pass@K Improvement",    "≥ 15% over baseline",         "P11 evaluation framework",     "Week 34"],
      ["Judge Score Improvement","≥ 0.5 point average increase","P11 LLM-as-judge pipeline",   "Week 34"],
      ["VRAM Usage",            "≤ 6GB peak during training",  "nvidia-smi or torch profiler", "Week 33"],
      ["LoRA Param Fraction",   "< 1% of total parameters",    "trainable_params / total_params","Week 32"],
    ],
    warning: "Dataset quality dominates fine-tuning outcome quality. 200 high-quality, diverse, instruction-following examples will outperform 2,000 noisy or repetitive examples. Invest significant time in Week 32 on data curation — remove length outliers (< 50 tokens or > 1,500 tokens), deduplicate aggressively, and manually review 50 random samples before training.",
    tip: "Run the P11 evaluation suite on the base model before any fine-tuning to establish a clean baseline. Without a pre-training benchmark, you cannot prove that fine-tuning actually helped. The two projects were deliberately designed to be used together — the evaluation framework is the measurement instrument for the fine-tuning experiment.",
  },

  // ── PROJECT 13 ─────────────────────────────────────────────────────────────
  {
    id: "P13", color: C.p13, phase: "V", weeks: "Weeks 35–37",
    title: "GPU SRAM Flash-Attention Simulator",
    subtitle: "Tiling Architecture · HBM vs. SRAM IO Analysis · Quadratic Memory Prevention",
    bottleneck: "Resolves memory-bound latency forced when the massive N×N intermediate attention score matrix must be written to global HBM memory during the standard softmax computation. In standard multi-head attention, computing attention scores for a sequence of length N produces an N×N matrix that must be materialised in GPU global memory (HBM). At sequence length 4,096, this single matrix requires 128MB of HBM per attention head. Flash-Attention eliminates this materialisation by restructuring the computation into small tiles that fit entirely in fast SRAM, performing the softmax reduction online.",
    overview: "This project implements a numerical simulator of the Flash-Attention algorithm using NumPy, allowing detailed instrumentation of memory read/write operations that would not be directly observable on real GPU hardware. The simulator implements the standard attention algorithm as a baseline, instruments all HBM read/write events, then implements the tiled Flash-Attention algorithm and demonstrates the memory IO reduction. A dashboard charts HBM access count vs. sequence length for both algorithms, empirically demonstrating Flash-Attention's linear IO complexity versus standard attention's quadratic IO complexity.",
    objectives: [
      "Implement standard scaled dot-product attention with full instrumented HBM IO tracking.",
      "Implement the Flash-Attention tiling algorithm with online softmax computation.",
      "Simulate SRAM block constraints: limit tile size to fit within a configurable SRAM budget.",
      "Chart HBM read/write operations vs. sequence length for both implementations.",
      "Validate numerical equivalence of Flash-Attention and standard attention outputs (within floating-point tolerance).",
    ],
    stack: {
      "Language": "Python 3.11",
      "Numerical Computing": "NumPy 1.26 (simulated attention, custom IO instrumentation)",
      "Visualisation": "Matplotlib (HBM IO count vs. sequence length, tile execution diagrams)",
      "Reference": "Flash-Attention paper (Dao et al., 2022) — read Sections 3–4 before implementing",
    },
    architecture: [
      "Standard Attention with IO Tracking: Implements Q×K^T → scale → mask → softmax → ×V as the standard sequence of matrix operations. A custom IO counter increments on every array read and write operation, tracking total HBM accesses. At sequence length N, standard attention performs O(N²) HBM reads/writes for the score matrix materialisation.",
      "Tiled Flash-Attention Algorithm: Divides Q, K, V into row tiles (Q blocks) and column tiles (K, V blocks). For each Q block, iterates over all K,V block pairs, computing partial attention scores and accumulating a running softmax numerator and denominator using the online log-sum-exp trick. The final output for each Q block is computed by combining running accumulators — the N×N score matrix is never materialised in HBM.",
      "Online Softmax Trick: Tracks the running maximum m_i and normalisation factor l_i across K,V tiles. When a new block introduces a larger maximum, the previous running output is rescaled by exp(m_old - m_new) before accumulation. This maintains numerical stability without requiring a two-pass computation over the full row.",
      "IO Complexity Analysis: The simulator logs HBM reads and writes independently for Q, K, V, attention scores, and output arrays. Standard attention IO scales as O(N²) due to score matrix materialisation. Flash-Attention IO scales as O(N) per tile × (N/tile_size) tiles = O(N) total, with a constant factor determined by SRAM tile size.",
    ],
    weekPlan: [
      { num: 35, title: "Standard Attention + IO Instrumentation", tasks: ["Implement standard scaled dot-product attention with full NumPy operations", "Build IO instrumentation layer: counter increments on every array read/write", "Run at sequence lengths N=128, 256, 512, 1024, 2048 — log total IO at each length", "Plot IO count vs. N: verify O(N²) growth curve", "Identify the score matrix materialisation step as the dominant IO contributor", "Annotate code with memory lifecycle comments: where data lives (HBM vs. SRAM conceptually)"], deliverable: "Instrumented standard attention implementation with O(N²) IO growth chart" },
      { num: 36, title: "Flash-Attention Tiling Implementation", tasks: ["Read Flash-Attention paper Sections 3–4: understand the tiling schema and online softmax", "Implement Q block outer loop, K/V block inner loop with configurable tile size", "Implement online softmax accumulation: running m_i, l_i, and output rescaling", "Validate: Flash-Attention output matches standard attention within 1e-5 absolute tolerance", "Run at same N sequence lengths — log IO count", "Plot Flash-Attention IO count vs. N on same chart as standard attention: observe linear vs. quadratic"], deliverable: "Flash-Attention implementation with validated numerical equivalence and IO comparison chart" },
      { num: 37, title: "Tile Size Analysis + Dashboard", tasks: ["Vary tile size (32, 64, 128, 256 elements) — measure IO count change", "Implement SRAM budget constraint: reject tile sizes exceeding configurable SRAM_SIZE parameter", "Chart HBM IO savings percentage across sequence lengths and tile size combinations", "Build summary dashboard: 4-panel figure (IO comparison, savings %, tile sensitivity, memory lifecycle)", "Write analysis: what SRAM tile size minimises IO for your simulated SRAM budget?", "Document: what sequence length threshold makes Flash-Attention break even vs. standard attention?"], deliverable: "4-panel dashboard with IO analysis, tile sensitivity study, and SRAM budget constraint implementation" },
    ],
    learnings: [
      "GPU storage layers: the hierarchy from SRAM (fast, small, on-chip) to HBM (slow, large, off-chip) and the latency implications of each.",
      "On-chip SRAM cache mechanics: why keeping intermediate computations in SRAM rather than writing them to HBM is the fundamental optimisation.",
      "IO complexity reduction via tiling: from O(N²) to O(N) by restructuring the computation order.",
      "Online log-sum-exp trick: numerically stable softmax computation without materialising the full score row.",
      "Memory-bound vs. compute-bound operations: why modern transformers are IO-limited, not FLOP-limited, on real GPU hardware.",
    ],
    metrics: [
      ["IO Complexity Validated",  "Flash: O(N), Std: O(N²)",    "Log-log slope of IO vs. N plot",  "Week 37"],
      ["Numerical Equivalence",    "< 1e-5 absolute difference",  "np.allclose comparison",          "Week 36"],
      ["IO Reduction at N=2048",   "≥ 80% fewer HBM accesses",   "Instrumented counter comparison", "Week 37"],
      ["Tile Sensitivity Tested",  "4 tile sizes profiled",       "Tile size sweep benchmark",       "Week 37"],
    ],
    warning: "Read the Flash-Attention paper (Dao et al., 2022) before writing any code in Week 36. The online softmax accumulation pattern — specifically the m_i and l_i running state update with rescaling — is non-obvious and will produce incorrect results if implemented from intuition alone. The paper's Algorithm 1 is the authoritative specification.",
    tip: "The key conceptual leap of Flash-Attention is that you can compute the correct softmax output without ever seeing all the scores simultaneously. The running max and normaliser allow you to 'correct' past outputs as new maximums are discovered. This is the insight that makes tiling possible — spend time building intuition for the online softmax before touching the tiling loop.",
  },

  // ── PROJECT 14 ─────────────────────────────────────────────────────────────
  {
    id: "P14", color: C.p14, phase: "V", weeks: "Weeks 38–40",
    title: "Heterogeneous Memory Layer & Model Offloader",
    subtitle: "CPU/VRAM Hybrid Execution · PCIe Pre-Fetching · Layer-by-Layer Offloading",
    bottleneck: "Addresses VRAM scarcity that blocks execution of multi-billion parameter models on consumer hardware. A 13B parameter model at FP16 requires approximately 26GB of VRAM — far beyond the 6–8GB in most consumer GPUs. Heterogeneous offloading shards the model across CPU RAM and GPU VRAM, executing layers on whichever compute unit holds their weights at a given step, with an asynchronous pre-fetcher streaming upcoming layers across the PCIe bus while current layers execute.",
    overview: "This project builds a custom layer-offloading engine using PyTorch hooks to intercept forward pass events. The engine partitions a multi-layer deep network across CPU RAM and GPU VRAM based on a configurable VRAM budget. During inference, it moves each layer to GPU immediately before execution and returns it to CPU immediately after, maintaining a fixed VRAM footprint regardless of total model size. An asynchronous pre-fetcher uses Python threads to stream the next layer to GPU during current layer execution, overlapping PCIe transfer latency with compute time.",
    objectives: [
      "Build a static layer partitioner that maps each model layer to CPU RAM or GPU VRAM based on a VRAM budget.",
      "Implement PyTorch forward hook interceptors for dynamic layer movement during inference.",
      "Build an asynchronous PCIe pre-fetching engine using Python threading.",
      "Benchmark: VRAM floor cap, PCIe transfer latency, and tokens-per-second across partitioning configurations.",
      "Validate: offloaded model produces identical outputs to a fully GPU-resident reference model.",
    ],
    stack: {
      "Language": "Python 3.11 (PyTorch hooks), C++ (optional, for custom CUDA/PCIe kernel)",
      "Framework": "PyTorch 2.2 (nn.Module hooks, torch.cuda.Stream for async transfers)",
      "Profiling": "torch.profiler, nvidia-smi (VRAM tracking), psutil (CPU RAM tracking)",
      "Concurrency": "Python threading (pre-fetch thread pool), asyncio for I/O coordination",
      "Target Model": "Multi-layer GPT-2 (124M) or P6 transformer for offloading experiments",
    },
    architecture: [
      "Static Layer Partitioner: Inspects each nn.Module layer's parameter byte count. Iterates through layers and assigns to GPU until the VRAM budget is exhausted; assigns remaining layers to CPU. Produces a partition map: {layer_id: 'cuda' | 'cpu'}. The partitioner is configurable with a VRAM_BUDGET_GB parameter.",
      "Forward Hook Interceptors: A pre_forward_hook attached to each CPU-resident layer calls layer.cuda() immediately before the layer's forward pass executes, moving its parameters to GPU. A post_forward_hook calls layer.cpu() after the forward pass completes, returning parameters to CPU RAM. This maintains a rolling 'active layer window' in VRAM.",
      "Asynchronous Pre-Fetcher: A background thread monitors which layer is currently executing. When layer i begins, the pre-fetcher initiates a non-blocking CUDA stream transfer of layer i+1 from CPU to GPU. The transfer runs on a separate CUDA stream, overlapping with layer i's compute. When layer i+1's hook fires, the transfer is complete with zero additional wait.",
      "VRAM Floor Cap: The minimum VRAM required by this engine equals the memory of the two largest consecutive layers (current layer + pre-fetched next layer), compared to the full model's VRAM requirement. This is the VRAM Floor Cap metric — the minimum hardware requirement for the model to execute at all.",
      "Inference Benchmark: Measures time-to-first-token and tokens-per-second for three configurations: fully GPU-resident (VRAM = full model size), partially offloaded with synchronous movement (baseline), and partially offloaded with asynchronous pre-fetching. The pre-fetcher's benefit is the latency difference between synchronous and asynchronous configurations.",
    ],
    weekPlan: [
      { num: 38, title: "Static Partitioning + Hook Interceptors", tasks: ["Inspect P6 or GPT-2 model: log layer-by-layer parameter count and memory footprint", "Build layer partitioner: assign layers to 'cuda' or 'cpu' based on VRAM_BUDGET_GB", "Implement pre_forward_hook: move CPU-resident layer to GPU before forward pass", "Implement post_forward_hook: return layer to CPU after forward pass completes", "Validate: offloaded model output matches fully GPU model output (torch.allclose)", "Benchmark: measure VRAM usage at peak during inference with hooks active"], deliverable: "Working offloading engine with PyTorch hooks and validated output equivalence" },
      { num: 39, title: "Sequential Offloading Benchmark", tasks: ["Measure synchronous per-layer transfer latency: CPU-to-GPU time for each layer (torch.cuda.Event timing)", "Plot layer transfer time vs. parameter count — establish PCIe bandwidth estimate", "Run full inference with synchronous offloading: measure tokens-per-second", "Identify bottleneck layers: largest parameter counts = longest PCIe transfers", "Experiment with VRAM budget levels: 2GB, 4GB, 6GB — measure tokens/sec at each", "Compare offloaded vs. fully GPU inference: quantify the offloading overhead cost"], deliverable: "Synchronous offloading benchmark with PCIe bandwidth measurement and overhead quantification" },
      { num: 40, title: "Async Pre-Fetcher + Final Dashboard", tasks: ["Implement pre-fetch thread: background CUDA stream transfer of layer i+1 during layer i execution", "Add synchronisation barrier: pre_forward_hook waits for pre-fetch stream completion before forward pass", "Benchmark: async vs. synchronous offloading tokens-per-second — measure overlap benefit", "Track VRAM Floor Cap: minimum peak VRAM across all configurations", "Build 4-panel summary dashboard: VRAM vs. config, tokens/sec vs. config, PCIe latency, layer timeline", "Write analysis: at what PCIe bandwidth does async pre-fetching become irrelevant?"], deliverable: "Async pre-fetching engine with 4-panel dashboard and VRAM Floor Cap analysis" },
    ],
    learnings: [
      "PCIe bus profiles: CPU-to-GPU transfer bandwidth (~16 GB/s for PCIe 4.0 x16) and how it determines offloading feasibility.",
      "Transfer overhead boundaries: minimum VRAM requirements and the point at which offloading becomes slower than not running the model at all.",
      "Asynchronous step planning loops: overlapping data transfer with compute using CUDA streams to hide PCIe latency.",
      "PyTorch hook system: pre_forward_hook and post_forward_hook as non-intrusive layer interception mechanisms.",
      "VRAM Floor Cap engineering: designing minimum viable hardware configurations for deploying models on constrained consumer hardware.",
    ],
    metrics: [
      ["VRAM Floor Cap",         "Measured per model configuration","Peak VRAM via nvidia-smi",     "Week 38"],
      ["PCIe Bandwidth",         "Estimated from transfer timings", "CUDA Event timing",             "Week 39"],
      ["Async Speedup",          "≥ 20% vs. synchronous baseline",  "Tokens/sec comparison",         "Week 40"],
      ["Output Equivalence",     "torch.allclose within 1e-5",      "Direct tensor comparison",      "Week 38"],
    ],
    warning: "PyTorch hooks fire on every forward pass call — including nested modules. If you attach hooks to an nn.Sequential container and also to its child modules, hooks fire multiple times per layer. Attach hooks exclusively to leaf modules (individual Linear, LayerNorm, etc.) and not to container modules (Sequential, ModuleList) to avoid double-transfers.",
    tip: "This project is the culmination of everything you have learned about memory hierarchies across the entire 40-week roadmap. The VRAM Floor Cap concept from Project 14 connects directly to the KV Cache from Project 8, the INT4 quantization from Project 8, and the tiling strategy from Project 13. Document these connections explicitly in your README — the synthesis is the proof of mastery.",
  },
];

// ─── BUILD DOCUMENT ───────────────────────────────────────────────────────────
const allSections = [];
const allChildren = [];

// ── COVER PAGE ────────────────────────────────────────────────────────────────
allChildren.push(
  spacer(1440),
  para("TECHNICAL ENGINEERING ROADMAP", { bold: true, color: C.navy, size: 20, align: AlignmentType.CENTER, before: 0, after: 200 }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 300 },
    children: [new TextRun({ text: "12-Month High-Performance AI Engineering", bold: true, color: C.navy, size: 52, font: "Arial" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.darkBlue, space: 6 } },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120 },
    children: [new TextRun({ text: "& Neural Systems Roadmap", bold: true, color: C.darkBlue, size: 48, font: "Arial" })],
  }),
  spacer(480),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 3, CONTENT_W / 3, CONTENT_W / 3],
    rows: [new TableRow({ children: [
      cell([para("14 Projects", { bold: true, color: C.white, size: 32, align: AlignmentType.CENTER, after: 20 }), para("Across 5 Phases", { color: "D0DCF0", size: 20, align: AlignmentType.CENTER, after: 0 })], { fill: C.navy, width: CONTENT_W/3 }),
      cell([para("40 Weeks", { bold: true, color: C.white, size: 32, align: AlignmentType.CENTER, after: 20 }), para("~10 Months", { color: "D0DCF0", size: 20, align: AlignmentType.CENTER, after: 0 })], { fill: C.darkBlue, width: CONTENT_W/3 }),
      cell([para("Full Stack to", { bold: true, color: C.white, size: 26, align: AlignmentType.CENTER, after: 20 }), para("Hardware Level", { color: "D0DCF0", size: 20, align: AlignmentType.CENTER, after: 0 })], { fill: C.midBlue, width: CONTENT_W/3 }),
    ]})]
  }),
  spacer(480),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: "Arveend  ·  Self-Learning Engineering Programme  ·  2025–2026", color: C.muted, size: 20, font: "Arial", italics: true })],
  }),
  new Paragraph({ children: [new PageBreak()] })
);

// ── TABLE OF CONTENTS PLACEHOLDER ────────────────────────────────────────────
allChildren.push(
  heading1("Table of Contents"),
  para("This document covers all 14 projects across 5 engineering phases. Each section contains: project overview, system bottleneck analysis, architectural design, week-by-week task breakdown, learning outcomes, success metrics, and critical implementation notes.", { size: 21, color: C.darkGray, after: 240 }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [7000, 2360],
    rows: [
      new TableRow({ children: [
        cell([para("SECTION", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 7000 }),
        cell([para("WEEKS", { bold: true, color: C.white, size: 19, after: 0, align: AlignmentType.CENTER })], { fill: C.navy, width: 2360 }),
      ]}),
      ...projects.map((p, i) => new TableRow({ children: [
        cell([para(`${p.id}  ·  ${p.title}`, { size: 19, color: i % 2 === 0 ? C.darkBlue : C.darkGray, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 7000 }),
        cell([para(p.weeks, { size: 19, color: C.muted, after: 0, align: AlignmentType.CENTER })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2360 }),
      ]}))
    ]
  }),
  spacer(240),
  new Paragraph({ children: [new PageBreak()] })
);

// ── PREPARATION ───────────────────────────────────────────────────────────────
allChildren.push(
  heading1("Environment Preparation Checklist"),
  para("The following environment configuration is mandatory before beginning any project. All 14 projects share a common infrastructure base. Establishing this foundation correctly in Week 0 prevents environment-related interruptions throughout the programme.", { size: 21, color: C.darkGray, after: 200 }),
  twoColTable(
    "Component", "Setup Requirement",
    ["Container Orchestration", "Local Model Deployment (P1–P5)", "Phase 2 Agent Models", "Monorepo Structure", "Financial API", "Load Testing", "Neural Computing", "Engineering Log"],
    ["Docker Desktop — manages PostgreSQL, Redis, Ollama containers", "Ollama: pull llama3.2:3b, qwen2.5-coder:7b, nomic-embed-text", "DeepSeek-r1 (Architect), Qwen2.5-coder (Developer), Llama3.1 (QA)", "Single GitHub monorepo with one folder per project", "Register Finnhub API key (free tier, ~60 calls/min)", "Install k6 v0.51+ for high-concurrency load testing", "Install PyTorch 2.2 and NumPy 1.26 for tensor operations (P6–P14)", "Obsidian or Notion: weekly log — Built / Broken / Resolved"],
    C.white, C.navy
  ),
  spacer(160),
  noteBox("CRITICAL", "Create a /mocks fixture directory in every project folder on Day 1. All external API responses (Finnhub, Ollama) must be mirrored to local JSON files on first invocation. This single practice prevents external rate limits and network availability from blocking development momentum across the entire 40-week timeline.", C.warning, C.warnBg),
  spacer(240),
  new Paragraph({ children: [new PageBreak()] })
);

// ── PHASE OVERVIEW ────────────────────────────────────────────────────────────
allChildren.push(
  heading1("Programme Architecture: Five Phases"),
  para("The 12-month programme is structured as five phases, each building directly on the engineering foundations of its predecessor. Projects within a phase may run in parallel where indicated. The Arveend Terminal (P4) is explicitly designed to integrate output from P1, P2, and P3.", { size: 21, color: C.darkGray, after: 200 }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [600, 1800, 2200, 1600, 3160],
    rows: [
      new TableRow({ children: [
        cell([para("Phase", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 600 }),
        cell([para("Title", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 1800 }),
        cell([para("Projects", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 2200 }),
        cell([para("Timeline", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 1600 }),
        cell([para("Core Focus", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 3160 }),
      ]}),
      ...phases.map((ph, i) => new TableRow({ children: [
        cell([para(ph.num, { bold: true, color: C.darkBlue, size: 20, after: 0, align: AlignmentType.CENTER })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 600 }),
        cell([para(ph.title, { bold: true, size: 19, color: C.black, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 1800 }),
        cell([para(projects.filter(p => p.phase === ph.num).map(p => p.id).join(", "), { size: 19, color: C.darkGray, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2200 }),
        cell([para(ph.weeks, { size: 19, color: C.muted, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 1600 }),
        cell([para(ph.desc, { size: 18, color: C.darkGray, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 3160 }),
      ]}))
    ]
  }),
  spacer(240),
  new Paragraph({ children: [new PageBreak()] })
);

// ── PROJECT SECTIONS ──────────────────────────────────────────────────────────
let currentPhase = null;
for (const proj of projects) {
  if (proj.phase !== currentPhase) {
    currentPhase = proj.phase;
    const ph = phases.find(p => p.num === proj.phase);
    allChildren.push(
      phaseBanner(ph.num, ph.title, ph.weeks, ph.desc),
      spacer(200),
    );
  }

  // Project banner
  allChildren.push(projectBanner(proj.id, proj.title, proj.weeks, proj.color));
  allChildren.push(spacer(60));

  // Subtitle
  allChildren.push(para(proj.subtitle, { size: 20, color: C.muted, italic: true, after: 160 }));

  // System Bottleneck
  allChildren.push(heading2("1. System Bottleneck Analysis", proj.color));
  allChildren.push(para(proj.bottleneck, { size: 21, color: C.darkGray, after: 160 }));

  // Overview
  allChildren.push(heading2("2. Project Overview", proj.color));
  allChildren.push(para(proj.overview, { size: 21, color: C.darkGray, after: 160 }));

  // Objectives
  allChildren.push(heading2("3. Engineering Objectives", proj.color));
  for (const obj of proj.objectives) allChildren.push(numItem(obj));
  allChildren.push(spacer(100));

  // Technology Stack
  allChildren.push(heading2("4. Technology Stack", proj.color));
  allChildren.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [2800, CONTENT_W - 2800],
    rows: [
      new TableRow({ children: [
        cell([para("Component", { bold: true, color: C.white, size: 19, after: 0 })], { fill: proj.color, width: 2800 }),
        cell([para("Technology", { bold: true, color: C.white, size: 19, after: 0 })], { fill: proj.color, width: CONTENT_W - 2800 }),
      ]}),
      ...Object.entries(proj.stack).map(([k, v], i) => new TableRow({ children: [
        cell([para(k, { bold: true, size: 19, color: C.darkBlue, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2800 }),
        cell([para(v, { size: 19, color: C.darkGray, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: CONTENT_W - 2800 }),
      ]}))
    ]
  }));
  allChildren.push(spacer(160));

  // Architecture
  allChildren.push(heading2("5. System Architecture", proj.color));
  for (const a of proj.architecture) {
    const parts = a.split(': ');
    if (parts.length >= 2) {
      allChildren.push(new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 40, after: 80 },
        children: [
          new TextRun({ text: parts[0] + ": ", bold: true, color: proj.color, size: 21, font: "Arial" }),
          new TextRun({ text: parts.slice(1).join(': '), size: 21, color: C.darkGray, font: "Arial" }),
        ]
      }));
    } else {
      allChildren.push(bulletItem(a));
    }
  }
  allChildren.push(spacer(120));

  // Week Plan
  allChildren.push(heading2("6. Week-by-Week Implementation Plan", proj.color));
  allChildren.push(weekTable(proj.weekPlan));
  allChildren.push(spacer(160));

  // Learning Outcomes
  allChildren.push(heading2("7. Learning Outcomes", proj.color));
  for (const l of proj.learnings) allChildren.push(bulletItem(l));
  allChildren.push(spacer(120));

  // Success Metrics
  allChildren.push(heading2("8. Success Metrics & Validation Criteria", proj.color));
  allChildren.push(metricsTable(proj.metrics));
  allChildren.push(spacer(160));

  // Warning + Tip
  allChildren.push(heading2("9. Implementation Notes", proj.color));
  allChildren.push(noteBox("CAUTION", proj.warning, C.warning, C.warnBg));
  allChildren.push(spacer(80));
  allChildren.push(noteBox("TIP", proj.tip, C.tip, C.tipBg));
  allChildren.push(spacer(240));

  allChildren.push(new Paragraph({ children: [new PageBreak()] }));
}

// ── CLOSING SUMMARY ───────────────────────────────────────────────────────────
allChildren.push(
  heading1("Programme Completion: Engineering Capability Matrix"),
  para("Upon completion of all 14 projects across 40 weeks, the following engineering competencies are demonstrably acquired through working, tested, and deployed software systems.", { size: 21, color: C.darkGray, after: 200 }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [2200, 3500, 3660],
    rows: [
      new TableRow({ children: [
        cell([para("Engineering Domain", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 2200 }),
        cell([para("Competency Acquired", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 3500 }),
        cell([para("Evidence Artifact", { bold: true, color: C.white, size: 19, after: 0 })], { fill: C.navy, width: 3660 }),
      ]}),
      ...[
        ["Backend Systems", "Distributed caching, atomic concurrency, async queues", "P2 k6 load chart, P3 chaos test assertions (P3)"],
        ["AI/RAG Systems", "Vector embedding, semantic retrieval, benchmarking", "P1 benchmark dashboard, P4 AI sentiment features"],
        ["LLM Orchestration", "Multi-agent pipelines, human-in-the-loop gates", "P5 visual canvas Loom demo recording"],
        ["Neural Architecture", "Transformer internals, attention, training loops", "P6 trained model with loss curves and perplexity"],
        ["Vector Indexing", "IVF clustering, HNSW graph search, ANN tradeoffs", "P7 QPS vs. recall benchmark analysis"],
        ["Quantization", "INT4 compression, KV cache, memory profiling", "P8 telemetry dashboard (4-panel)"],
        ["Advanced Inference", "MoE routing, speculative decoding, evaluation", "P9–P11 benchmark dashboards"],
        ["Fine-Tuning", "LoRA PEFT, adapter merging, 6GB VRAM constraint", "P12 training curves, P11 evaluation delta"],
        ["Hardware Optimisation", "Flash-Attention tiling, PCIe offloading, VRAM analysis", "P13–P14 IO analysis dashboards"],
      ].map(([a, b, c], i) => new TableRow({ children: [
        cell([para(a, { bold: true, size: 19, color: C.darkBlue, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 2200 }),
        cell([para(b, { size: 19, color: C.black, after: 0 })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 3500 }),
        cell([para(c, { size: 18, color: C.muted, after: 0, italic: true })], { fill: i % 2 === 0 ? C.veryLight : C.white, width: 3660 }),
      ]}))
    ]
  }),
  spacer(200),
  para("Every competency above is backed by a concrete, testable software artifact — not theoretical study notes. The portfolio produced by this programme represents demonstrable engineering capability across the full AI systems stack from infrastructure to hardware.", { size: 21, color: C.muted, italic: true, align: AlignmentType.CENTER, before: 120, after: 0 })
);

// ─── ASSEMBLE DOCUMENT ────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } }, run: { font: "Arial" } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 960, hanging: 240 } }, run: { font: "Arial" } } },
      ]},
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } }, run: { font: "Arial" } } },
      ]},
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: C.black } },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: C.heading },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: C.heading },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: C.darkBlue },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.darkBlue, space: 4 } },
          children: [new TextRun({ text: "12-Month AI Engineering Roadmap  ·  Arveend", color: C.muted, size: 18, font: "Arial", italics: true })],
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.darkBlue, space: 4 } },
          children: [
            new TextRun({ text: "Page ", color: C.muted, size: 18, font: "Arial" }),
            new TextRun({ text: " ", color: C.muted, size: 18, font: "Arial" }),
            new TextRun({ text: "  ·  Self-Learn Engineering Programme  ·  2025–2026", color: C.muted, size: 18, font: "Arial" }),
          ],
        })]
      })
    },
    children: allChildren,
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/ap/quiz/AI_Engineering_Roadmap.docx", buffer);
  console.log("Done. Saved to /Users/ap/quiz/AI_Engineering_Roadmap.docx");
});