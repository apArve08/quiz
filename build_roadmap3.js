'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
} = require('docx');
const fs = require('fs');

const CW = 9360;
const bdr = (c='BBBBCC') => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const allB = (c='BBBBCC') => ({ top:bdr(c), bottom:bdr(c), left:bdr(c), right:bdr(c) });

function cell(children, opts={}) {
  return new TableCell({
    children,
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    borders: allB(opts.bc || 'BBBBCC'),
  });
}
function p(text, opts={}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before||0, after: opts.after||100, line: 276 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: [new TextRun({ text, bold:opts.bold||false, italics:opts.italic||false, color:opts.color||'1A1A2E', size:opts.size||22, font:'Arial' })],
    border: opts.border || undefined,
  });
}
const h1 = (text) => new Paragraph({ heading:HeadingLevel.HEADING_1, spacing:{before:320,after:160}, children:[new TextRun({text,bold:true,color:'1B2A4A',size:36,font:'Arial'})], border:{bottom:{style:BorderStyle.SINGLE,size:8,color:'2E5090',space:4}} });
const h2 = (text, color='2E5090') => new Paragraph({ heading:HeadingLevel.HEADING_2, spacing:{before:240,after:120}, children:[new TextRun({text,bold:true,color,size:26,font:'Arial'})] });
const sp = (n=120) => new Paragraph({ spacing:{before:0,after:n}, children:[new TextRun('')] });
const pb = () => new Paragraph({ children:[new PageBreak()] });
const bullet = (label, text) => new Paragraph({ numbering:{reference:'bullets',level:0}, spacing:{before:40,after:70}, children: label ? [new TextRun({text:label+': ',bold:true,color:'2E5090',size:21,font:'Arial'}),new TextRun({text,size:21,font:'Arial',color:'333355'})] : [new TextRun({text,size:21,font:'Arial',color:'333355'})] });
const num = (text) => new Paragraph({ numbering:{reference:'numbers',level:0}, spacing:{before:40,after:70}, children:[new TextRun({text,size:21,font:'Arial',color:'333355'})] });
function tbl(rows, colWidths) { return new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:colWidths, rows }); }
const row = (cells) => new TableRow({ children: cells });
function noteBox(label, text, labelColor, textColor, bg, borderColor) {
  return tbl([row([
    cell([p(label,{bold:true,color:labelColor,size:19,after:0,align:AlignmentType.CENTER})],{fill:bg,w:1200,bc:borderColor}),
    cell([p(text,{size:19,color:textColor,after:0})],{fill:bg,w:CW-1200,bc:borderColor}),
  ])],[1200,CW-1200]);
}

const projects = [
  { id:'P1', weeks:'Wk 1–4', color:'4B3F9E', light:'EAE8FF', title:'Local AI Evaluator & RAG Pipeline', phase:'Phase I  ·  Backend Infrastructure',
    focus:'Vector Math · Embeddings · pgvector · LangChain · Ollama · Next.js',
    why:'Cloud LLM APIs are expensive, rate-limited, and leak sensitive data. This teaches you to run your own local AI pipeline — the same RAG architecture inside Notion AI, Cursor, and GitHub Copilot — built from scratch so you understand every layer.',
    what:['Document ingestion engine: load PDFs and text files, chunk into overlapping segments','Embedding generation: convert each chunk into a 768-dimension vector using nomic-embed-text via Ollama','pgvector database: store and similarity-search vectors in PostgreSQL with the cosine distance operator','Benchmarking harness: loop 20 test questions through 2 local LLMs, log TTFT and tokens per second','Next.js analytics dashboard: latency charts, accuracy scatter plots, live question tester'],
    learn:['Embeddings: text represented as a point in 768-dimensional space where similar meaning equals close distance','Cosine similarity vs L2 distance: when each distance metric is appropriate for retrieval tasks','Chunking strategy: how chunk size (256 vs 1024 tokens) and overlap percentage affect retrieval quality','The full RAG loop: retrieve relevant chunks, augment the prompt, generate the grounded response','Model benchmarking: TTFT, tokens per second, manual accuracy — measuring LLM performance objectively','pgvector operators: cosine, L2, and inner product distance and when to apply each'],
    weekplan:[
      {wk:'1',title:'Ingestion & Chunking',task:'Install Ollama, pull llama3.2:3b, deepseek-r1:7b, nomic-embed-text. Write Python ingest script with LangChain RecursiveCharacterTextSplitter. Test chunk sizes 256, 512, 1024. Log chunk counts. Pin requirements.txt on Day 1.'},
      {wk:'2',title:'Vector Layer',task:'Launch PostgreSQL with pgvector in Docker. Create embeddings table with vector(768) column. Batch-embed all chunks using nomic-embed-text. Run a manual similarity query: top-5 chunks most similar to a test sentence.'},
      {wk:'3',title:'Benchmark Core',task:'Write 20 test questions at varying difficulty. Build loop: question to cosine similarity query to top-3 chunks to augmented prompt to both models simultaneously. Log TTFT, tokens per second, manual accuracy (1 to 5) to JSON file.'},
      {wk:'4',title:'Analytics Dashboard',task:'Scaffold Next.js 14 with Recharts. Build 3 charts: latency bar, tokens per second line, accuracy scatter plot. Add live Run Test input that fires benchmark and appends results to all charts dynamically.'},
    ],
    hardPart:'Chunking is 80% of RAG quality. A bad chunk size produces irrelevant retrievals even with a perfect model. Experiment in Week 1 before touching the vector store.',
    tip:'Use your own documents — RunTrack session notes or GastroHub schemas. Real data makes benchmark results personally meaningful and domain-specific.',
    deliverable:'Next.js benchmark dashboard + GitHub repo with Docker one-command startup' },

  { id:'P2', weeks:'Wk 5–8', color:'9E3F3F', light:'FFE8E8', title:'Distributed Cache-Aside System', phase:'Phase I  ·  Backend Infrastructure',
    focus:'Redis · Cache-Aside Pattern · TTL Management · Dirty Read Prevention · k6 Load Testing',
    why:'Every database has a throughput ceiling. Without a caching layer, a moderate traffic spike will pin your PostgreSQL CPU to 100% and cascade into timeouts. This is the exact architecture Shopify and Stripe use to serve millions of reads without hitting the database.',
    what:['Two-tier data layer: PostgreSQL as source of truth plus Redis in-memory cache via Docker Compose','Cache-Aside GET middleware: check Redis first, on miss fetch PostgreSQL and populate cache with TTL','Atomic write path: PUT and PATCH wraps SQL UPDATE plus Redis DEL in one coherent sequence','Cache stampede protection: Redis SETNX mutex prevents 100 concurrent misses hitting PostgreSQL at once','k6 load test: 3,000 virtual users comparing DB CPU% with versus without the caching layer'],
    learn:['Cache-Aside vs Write-Through vs Read-Through: three caching patterns and when each applies','TTL tradeoffs: too short means constant cache misses, too long means stale data served to users','Dirty read problem: why cache must be invalidated atomically with the write, never before or after','Cache stampede: the failure mode where a cache miss multiplies database load instead of reducing it','k6 load testing: virtual user ramping, p95 and p99 latency metrics, threshold configuration'],
    weekplan:[
      {wk:'5',title:'Two-Tier Setup',task:'Docker Compose: PostgreSQL plus Redis with named volumes. Design schema: news_items with 10+ columns. Seed 20,000 records with faker.js. Build baseline Express TypeScript CRUD with zero caching. Measure and log baseline GET latency average, p95, and p99.'},
      {wk:'6',title:'Cache-Aside Middleware',task:'Install ioredis. Implement GET middleware: Redis check, hit returns immediately, miss fetches PostgreSQL then SET Redis with TTL 60 seconds. Add X-Cache HIT or MISS response header. Confirm second call under 2ms. Log hit and miss ratio via Redis INCR.'},
      {wk:'7',title:'Atomic Invalidation',task:'PUT and PATCH: wrap SQL UPDATE plus Redis DEL in one sequence. Write dirty-read test: update record, immediate GET, assert fresh data returned. Implement SETNX mutex for stampede protection. Write integration test: create, cache read, update, fresh read, delete.'},
      {wk:'8',title:'Load Test Proof',task:'Write k6 script: ramp 0 to 3,000 VUs over 60 seconds. Run with caching DISABLED, capture PostgreSQL CPU%. Run with caching ENABLED, capture same metrics. Build comparison chart. This DB CPU comparison screenshot is your primary portfolio proof artifact.'},
    ],
    hardPart:'Cache stampede: when 100 concurrent requests miss the same key, they all query PostgreSQL simultaneously — multiplying load exactly when you are trying to reduce it. Implement SETNX locking before any load test.',
    tip:'Your MES background maps directly here. Production floor data refresh faces the same TTL design problem: too short hammers the DB, too long shows operators stale machine states.',
    deliverable:'Express API with k6 load test proving 70% or greater DB CPU reduction, deployed to Railway' },

  { id:'P3', weeks:'Wk 9–12', color:'3F7A9E', light:'E8F0FF', title:'High-Performance Inventory Reservation Engine', phase:'Phase I  ·  Backend Infrastructure',
    focus:'Redis Lua Scripts · BullMQ · Atomic Concurrency · Chaos Testing · Eventual Consistency',
    why:'Overselling is catastrophic in any booking or e-commerce system. The root cause is always a non-atomic read-check-write across concurrent threads. This teaches you the Redis single-threaded model and how Lua scripts make race conditions physically impossible.',
    what:['Redis stock counters: seed in-memory product stock at product:stock:{id} = 100 per product','Intentionally broken endpoint: implement WITHOUT atomic protection and empirically prove the race condition','Lua atomic script: IF stock > 0 THEN DECRBY 1 RETURN success — one single uninterruptible operation','BullMQ pipeline: successful checkouts enqueue to Redis queue, worker persists to PostgreSQL asynchronously','Chaos test: 2,000 concurrent requests for 5 units — assert exactly 5 successes and 1,995 rejections'],
    learn:['Why race conditions happen: the interleaving window between read and write in non-atomic operations','Redis single-threaded model: why it makes Lua scripts an atomic primitive that cannot be interrupted','EVALSHA pattern: pre-loading Lua scripts via SCRIPT LOAD for performance without re-transmission overhead','BullMQ job queue: producer to queue to consumer worker with dependency chains and retry logic','Eventual consistency: checkout responds in microseconds, PostgreSQL write happens in the background','Chaos engineering: writing tests with explicit invariant assertions rather than just checking absence of errors'],
    weekplan:[
      {wk:'9',title:'Broken First',task:'Build checkout WITHOUT atomic protection: Redis GET, check, DECR as three separate calls. Run 500 concurrent axios requests. Screenshot negative stock and wrong success counts. Draw the thread interleaving timeline. This before state makes the fix meaningful.'},
      {wk:'10',title:'Lua Atomic Lock',task:'Write Lua script: IF stock greater than 0 THEN DECRBY 1 RETURN 1 ELSE RETURN 0. Load via SCRIPT LOAD, call via EVALSHA. Re-run 500 concurrent test — verify zero overselling. Add 10-minute reservation expiry with EXPIRE command.'},
      {wk:'11',title:'Event Pipeline',task:'Install BullMQ v5. On success: push order payload to queue. Worker writes invoice to PostgreSQL with idempotency guard. Add 3x retry with exponential backoff at 1s, 4s, 16s. Route permanent failures to dead-letter queue. Deploy BullMQ Board UI.'},
      {wk:'12',title:'Chaos Script',task:'Write chaos test: 2,000 concurrent requests, product with 5 units. Assert 4 invariants: exactly 5 HTTP 200 responses, exactly 1,995 HTTP 409 responses, Redis stock equals 0, PostgreSQL invoice count equals 5. Run 3 times — deterministic result each time. Screenshot proof.'},
    ],
    hardPart:'Lua scripts cannot call external services or do async I/O in Redis. Keep the script purely computational: read, evaluate, write, return. Any external call causes a Redis error and breaks the atomic guarantee.',
    tip:'This is your strongest portfolio artifact. A Loom of the chaos test showing 5 successes and 1,995 clean rejections makes senior engineers stop and look in technical interviews.',
    deliverable:'Atomic checkout with chaos test showing 4 passing invariants across 3 deterministic runs' },

  { id:'P4', weeks:'Wk 1–12 (Parallel)', color:'9E7A1A', light:'FFF8E0', title:'Arveend Fintech Terminal', phase:'Phase I  ·  Backend Infrastructure (Capstone Integration)',
    focus:'Next.js 14 · Socket.io WebSockets · Finnhub API · Redis Cache · Paper Trading Wallet',
    why:'P4 is the capstone integration of P1, P2, and P3. RAG for earnings summaries from P1, Cache-Aside for stock profiles from P2, and atomic reservation for paper trading orders from P3 all wire together here into one live product.',
    what:['V1 Weeks 1–4: Stock search, company profiles, candlestick charts, portfolio P&L tracker with JWT auth','V2 Weeks 5–8: Ollama sentiment scoring on news, RAG earnings summarizer, market screener with filters','V3 Weeks 9–12: WebSocket price feed at 500ms ticks, Redis alert processor, paper trading wallet','Mock fixture layer: mirror all Finnhub API responses to local JSON on first call — develop without rate limits','Deploy V1 to Railway by end of Week 4 — live public URL established from the very start'],
    learn:['WebSocket lifecycle: subscription rooms, reconnection strategies, server-side broadcasting to multiple clients','Reactive UI optimisation: circular buffer for tick history, imperative chart API to avoid full component re-renders','Portfolio accounting: cost basis computation, unrealised P&L calculation, position reconciliation','Third-party API rate limit management: fixture-layer development, TTL-based caching, graceful degradation','Full-stack integration: applying P1 RAG, P2 Cache-Aside, P3 atomic reservation in one deployed product'],
    weekplan:[
      {wk:'V1  Wk 1–4',title:'Core Engine',task:'Finnhub integration plus mock fixture layer. PostgreSQL schema: users, watchlist, transactions. JWT auth middleware. SQL P&L aggregation function. Next.js candlestick chart using lightweight-charts. Deploy to Railway by Week 4 — live URL from this point forward.'},
      {wk:'V2  Wk 5–8',title:'AI Intelligence',task:'Ollama sentiment: 10 articles to llama3.2:3b to JSON output with score, label, and reasoning fields. RAG earnings summarizer using P1 pgvector database. Backend math: beta, moving averages, dividend yield. Screener data table with PE ratio, yield, and beta filters.'},
      {wk:'V3  Wk 9–12',title:'Streaming & Trading',task:'Socket.io server: emit price ticks every 500ms per subscribed symbol. Client circular buffer: append ticks without full component re-render. Redis alert processor: compare tick to user target price, emit alert event. Paper trading wallet using P3 Lua atomic pattern for order execution.'},
    ],
    hardPart:"Finnhub's free tier is approximately 60 API calls per minute. Without a mock fixture layer you exhaust the daily limit within 20 minutes of active development. Build the fixture layer on Day 1 of Week 1 — not when you hit the rate limit.",
    tip:'Deploy V1 to Railway by Week 4 even if only charts work. A live URL with real data is worth more for your portfolio than a perfect local application with no public presence.',
    deliverable:'Live Arveend Terminal with V1 plus V2 plus V3 deployed, WebSocket streaming, AI sentiment, and paper trading' },

  { id:'P5', weeks:'Wk 13–15', color:'5A3F9E', light:'EDE8FF', title:'Visual Multi-Agent Workspace', phase:'Phase II  ·  Agentic Orchestration',
    focus:'Multi-LLM Orchestration · BullMQ Dependency Chains · Human-in-the-Loop · Socket.io Canvas',
    why:'Single-agent AI cannot parallelise specialist work and cannot be supervised at meaningful checkpoints. This teaches the orchestration patterns behind AutoGPT, CrewAI, and enterprise AI pipelines — built with full understanding of every queue event and socket message.',
    what:['Three role-assigned agents: Architect via DeepSeek-r1, Developer via Qwen2.5-coder, QA via Llama3.1','BullMQ dependency chains: Developer job is blocked until Architect job completes and is approved','Human-in-the-loop approval gates: pipeline pauses at checkpoints with Approve and Reject with Feedback buttons','Next.js visual canvas: swimlane layout showing live agent status across Queued, Thinking, Writing, Awaiting Review','Streaming tokens: each agent streams via Ollama through Socket.io directly into the live UI'],
    learn:['Multi-agent context management: each model has its own isolated context window with no shared state','BullMQ job dependency resolution: child jobs are blocked until parent jobs complete successfully','Human-in-the-loop governance: designing approval gates that balance automation speed with human oversight','Streaming buffer arrays: handling partial token delivery and rendering incremental output without re-renders','Producer-consumer architecture: orchestrator as producer, agent workers as typed job consumers'],
    weekplan:[
      {wk:'13',title:'Orchestrator Foundation',task:'Multi-model Ollama environment: all three models available simultaneously. Task decomposition endpoint: high-level prompt returns JSON subtask array. BullMQ dependency chain configuration. Agent worker stubs. Socket.io server for status events.'},
      {wk:'14',title:'Agents + Canvas',task:'Architect worker: DeepSeek-r1 streams tokens to Socket.io. Developer worker: Qwen2.5-coder with architecture output as system context. QA worker: Llama3.1 reviews Developer output. Next.js swimlane canvas. Human approval gate: Approve or Reject with Feedback.'},
      {wk:'15',title:'E2E Test + Docs',task:'Submit real engineering task: build a REST CRUD API. Measure full pipeline execution time. Add task history to PostgreSQL. Write README with agent role diagram. Record Loom showing all three agents working simultaneously on screen.'},
    ],
    hardPart:'Running three large models simultaneously requires significant RAM, ideally 32GB or more. If memory is constrained, run agents sequentially — the architectural pattern and lessons are identical.',
    tip:'Your MES workflow of operator to engineer to QA sign-off is structurally identical to this project. You already understand why each approval gate exists. You are now implementing it in software.',
    deliverable:'Visual multi-agent canvas with human approval gates and end-to-end pipeline test recording' },

  { id:'P6', weeks:'Wk 16–18', color:'9E4A3F', light:'FFE8E5', title:'Transformer Architecture from Scratch', phase:'Phase III  ·  Neural Systems Engineering',
    focus:'PyTorch Raw Tensors · Q/K/V Attention Matrices · Causal Masking · BPE Tokenizer · AdamW Training',
    why:'Engineers who only use Hugging Face cannot debug attention anomalies, cannot reason about gradient flow, and cannot make informed architecture decisions. Building a transformer from raw tensors gives you the mathematical intuition that separates engineers who understand the system from those who just use it.',
    what:['Custom BPE tokenizer: iterative merge algorithm from scratch with no tiktoken or sentencepiece dependency','Token and positional embeddings: learned tables mapping integer token IDs to d_model-dimensional vectors','Multi-head causal attention: Q/K/V projections, score matrix, causal mask, softmax, output projection','Feedforward block: 2-layer MLP with GELU activation and Pre-LayerNorm for training stability','Training loop: AdamW with gradient clipping at norm 1.0, cross-entropy loss, linear LR warmup schedule'],
    learn:['Q, K, V matrices: queries search keys to retrieve weighted values — why this formulation enables attention','Causal masking: the upper-triangular negative infinity mask enforces the autoregressive constraint during training','Multi-head splitting: parallel heads learn different relationship types in the same sequence simultaneously','Residual connections and LayerNorm: why deep networks need skip connections to propagate gradients effectively','Cross-entropy loss: the mathematical relationship between predicted logit distributions and the training signal','AdamW vs Adam: weight decay decoupling and why it matters for transformer training stability'],
    weekplan:[
      {wk:'16',title:'Tokenizer + Embeddings',task:'Character-level tokenizer: build vocabulary, encode and decode functions. BPE tokenizer: iterative merge algorithm with configurable vocabulary size. Token embedding table. Sinusoidal positional encoding. Unit tests: encode to decode round-trip. Corpus: Shakespeare complete works at approximately 1MB.'},
      {wk:'17',title:'Attention + Transformer Block',task:'Scaled dot-product attention: Q/K/V projections, score equals QK-transpose divided by sqrt(d_k), causal mask as additive negative infinity, softmax, output projection. Multi-head: parallel heads, concatenate, project. Full decoder block: Attention plus FFN plus residuals plus LayerNorm. Stack N blocks. Validate forward pass output shape.'},
      {wk:'18',title:'Training Loop + Evaluation',task:'AdamW loop: forward, cross-entropy loss, backward, gradient clip at norm 1.0, parameter update. Linear LR warmup over 100 steps. Train 2,000 steps. Plot loss curve. Generate sample text at checkpoints. Visualise attention heat maps. Report validation perplexity on held-out set.'},
    ],
    hardPart:'Do NOT use nn.MultiheadAttention or nn.TransformerDecoderLayer. These wrapper modules hide the math you need to understand. Implement attention as raw torch.matmul operations with manual reshape and permute for multi-head splitting.',
    tip:'Train on Shakespeare. It is small at about 1MB, trains in minutes on CPU, and produces text samples interesting enough to qualitatively evaluate. Resist the temptation to use a larger dataset — the goal is understanding, not capability.',
    deliverable:'Trained transformer with loss curve, attention heat maps, perplexity score, and generated text samples' },

  { id:'P7', weeks:'Wk 19–21', color:'3F9E6A', light:'E8FFF3', title:'Vector Indexing Engine', phase:'Phase III  ·  Neural Systems Engineering',
    focus:'IVF k-means Clustering · HNSW Graph Search · ANN Algorithms · NumPy · O(log N) Complexity',
    why:'pgvector uses HNSW and IVF under the hood. Without understanding these algorithms, you cannot tune index parameters, reason about recall versus speed tradeoffs, or choose between index types. This project builds both algorithms from raw NumPy.',
    what:['Brute-force linear scan: exact nearest neighbour as ground truth for recall measurement','IVF index: k-means clustering using Lloyds algorithm, Voronoi cell partitioning, inverted lists per centroid','IVF query: compute distances to all centroids, select nprobe nearest, scan their inverted lists, return top-K','HNSW index: hierarchical proximity graph with bidirectional edges and multi-layer insertion','HNSW query: top-layer greedy descent through layers then bottom-layer beam search with ef parameter'],
    learn:['Curse of dimensionality: why Euclidean distance loses discriminative power at 768 or more dimensions','Voronoi decomposition: how k-means partitions vector space into cells for approximate nearest-neighbour search','HNSW small-world graph properties: why hierarchical routing achieves O(log N) average complexity','Recall versus speed tradeoff: nprobe in IVF and ef in HNSW as configurable accuracy control parameters','Build-time versus query-time complexity: IVF builds fast, HNSW queries faster at high recall targets'],
    weekplan:[
      {wk:'19',title:'Brute-Force + IVF Build',task:'Vectorised cosine baseline using NumPy matrix operations. Lloyds k-means: random initialisation, assignment step, centroid update, convergence check. IVF: assign corpus vectors to nearest centroids, build inverted lists. IVF query: centroid distance ranking, nprobe selection, candidate set search. Benchmark QPS and recall at 10.'},
      {wk:'20',title:'HNSW Implementation',task:'HNSW node and layer data structures in Python dictionaries. Insert: random layer assignment, greedy entry search, M-nearest-neighbour bidirectional edges. Query: top-layer entry, layer-by-layer greedy descent, bottom-layer beam search. Validate recall at 10 above 95 percent. Full benchmark: HNSW vs IVF vs brute-force.'},
      {wk:'21',title:'Visualisation + Analysis',task:'Visualise IVF Voronoi cells using 2D PCA projection of real embedding data. Visualise HNSW layer graph for a 50-node toy dataset. Plot QPS vs recall curves for both algorithms across parameter sweeps. Write analytical report on when to choose IVF vs HNSW vs pgvector. Package as importable Python library.'},
    ],
    hardPart:'Do NOT use faiss, annoy, or hnswlib. Those are exactly the libraries you are reimplementing. All vector arithmetic must be raw NumPy operations. Using a library defeats the purpose of this project entirely.',
    tip:'Start IVF with k=10 centroids on 1,000 vectors before scaling to 100K. Visualise Voronoi cells on 2D synthetic data to build spatial intuition before benchmarking numerically.',
    deliverable:'Python ANN library with IVF and HNSW, benchmark comparison chart, and analytical report' },

  { id:'P8', weeks:'Wk 22–25', color:'9E3F7A', light:'FFE8F5', title:'Quantization & Model Inference Engine', phase:'Phase III  ·  Neural Systems Engineering',
    focus:'FP32 to INT4 Weight Compression · KV Cache Architecture · Memory Profiling · C++ Dequant Kernel',
    why:'A 7B parameter model at FP32 requires 28GB VRAM — impossible on consumer hardware. Quantization is what makes local LLMs possible. Understanding it at bit level lets you predict accuracy loss, choose the right compression scheme, and design systems within hardware constraints.',
    what:['Symmetric INT8 quantization: scale equals max absolute W divided by 127, quantize and dequantize, measure size reduction','Asymmetric INT4 with group-wise calibration: group size 128, separate scale and zero-point per group','Per-channel calibration pipeline: forward pass on representative data to set activation-aware scale factors','KV Cache: pre-allocated tensor buffer [batch, n_layers, n_heads, max_seq, head_dim] for K/V reuse across steps','C++ INT4 dequantization kernel: pack two 4-bit weights per byte, callable from Python via ctypes'],
    learn:['Quantization arithmetic: how FP32 maps to INT8 and INT4 using scale and zero-point, where rounding error accumulates','Memory-bandwidth-bound inference: why LLMs on consumer GPUs are DRAM-limited, not FLOP-limited','KV Cache mechanics: why recomputing prefix attention at every generation step is redundant and costly','Group-wise vs per-tensor quantization granularity: the accuracy-compression tradeoff and when each applies','INT4 bit packing: storing two 4-bit values per byte and the bit manipulation required to unpack them'],
    weekplan:[
      {wk:'22',title:'FP32 to INT8',task:'Symmetric INT8: scale computation, quantize and dequantize functions. Apply to all linear layers in P6 transformer model. Measure model size in MB before and after. Track memory usage with tracemalloc. Evaluate on 20 P1 benchmark questions: INT8 vs FP32 accuracy comparison.'},
      {wk:'23',title:'INT4 + Calibration',task:'Asymmetric INT4 with group size 128: scale and zero-point per group. Per-channel calibration pipeline. C++ dequantization function callable via Python ctypes. Three-way comparison: FP32 vs INT8 vs INT4 across size in MB, perplexity, and generation quality.'},
      {wk:'24',title:'KV Cache',task:'Pre-allocated KV buffer tensor. Modify P6 forward pass: write K and V projections to cache at each generation step. Concatenate cached K,V with new token projections before attention. Measure tokens per second with and without KV Cache at prefix length 512. Validate identical output.'},
      {wk:'25',title:'Telemetry Dashboard',task:'4-panel Matplotlib dashboard: model size, memory usage, perplexity, tokens per second versus precision level. Compression ratio vs accuracy degradation curve. Analysis report on compute-bound versus memory-bound inference. Package as reusable Python module for use in P12.'},
    ],
    hardPart:'INT4 quantization of a poorly trained base model compounds errors into incoherent output. If the P6 transformer quality is too low, use Hugging Face GPT-2 (124M) as your target model — this is explicitly permitted for this project.',
    tip:'The KV Cache component is the most practically valuable part of this project. In production inference servers like vLLM and TensorRT-LLM, KV Cache management is the primary throughput bottleneck. Understanding it from scratch gives you rare depth.',
    deliverable:'Quantization module (INT8 and INT4) plus KV Cache plus 4-panel telemetry dashboard' },

  { id:'P9', weeks:'Wk 26–27', color:'6A9E3F', light:'F0FFE8', title:'Sparse Mixture of Experts (MoE) Transformer', phase:'Phase IV  ·  Advanced Inference & MLOps',
    focus:'Top-K Gating Router · Expert Networks · Auxiliary Load Balancing Loss · Routing Collapse Prevention',
    why:'MoE is the architecture behind GPT-4, Mixtral, and DeepSeek-V3. Understanding sparse routing is the difference between knowing how to use these models and understanding why they are computationally efficient — and why routing collapse is their most dangerous failure mode.',
    what:['N=8 independent expert FFN modules implemented as separate nn.Module instances','Gating network: linear projection from d_model to N experts, Softmax, top-K=2 selection per token','Auxiliary Load Balancing Loss: penalises routing collapse with alpha coefficient controlling strength','Token capacity constraint: maximum tokens per expert per batch — overflow tokens are dropped','Ablation experiment: train without Auxiliary Loss to observe routing collapse, then fix it'],
    learn:['Sparse activation: only K of N experts fire per token, reducing FLOPs while maintaining total model capacity','Auxiliary Loss formula: dot product of selection frequency and routing probability enforces uniform expert balance','Routing collapse: why models learn to ignore most experts and how the alpha coefficient prevents this failure mode','Softmax gate temperature: the relationship between logit scaling and routing sharpness','Expert capacity constraints: why token dropping is preferred over dynamic batching overhead in distributed settings'],
    weekplan:[
      {wk:'26',title:'MoE Layer',task:'Build N=8 expert FFN modules as independent nn.Module instances. Gating network: linear to N, Softmax, top-K=2 selection, renormalise selected weights. Weighted sum of K expert outputs per token. Auxiliary Load Balancing Loss with alpha=0.01. Unit test: exactly K experts activate per token, all others output zero.'},
      {wk:'27',title:'Integration + Analysis',task:'Replace P6 feedforward block with MoE layer as a drop-in swap. Train 500 steps: verify loss decreases and aux loss converges. Compare MoE vs dense feedforward loss curves. Plot expert utilisation bar chart across 1,000 training batches. Ablation: set alpha to zero and observe routing collapse within 50 steps.'},
    ],
    hardPart:'Routing collapse happens fast — often within 50 training steps without Auxiliary Loss. Start with alpha=0.01 and increase to 0.1 if collapse is still observed. Plot expert utilisation every 100 steps during training.',
    tip:'Ground yourself in the simple case first: N=4 experts, K=1 top-1 routing, 200 training steps. Once Auxiliary Loss works at small scale, expand to N=8 and K=2.',
    deliverable:'MoE transformer with uniform expert utilisation chart, routing collapse ablation, and activated-parameter comparison' },

  { id:'P10', weeks:'Wk 28–29', color:'3F6A9E', light:'E8F0FF', title:'Speculative Decoding Engine', phase:'Phase IV  ·  Advanced Inference & MLOps',
    focus:'Draft-Verify Loop · Residual Acceptance Sampling · Throughput Benchmarking · Dual-Model Inference',
    why:'Sequential generation reloads all model weights for every single token — that is the memory bandwidth bottleneck. Speculative decoding is how vLLM and TGI achieve 2 to 4x throughput improvement without changing the model. The acceptance sampling math proves the output distribution is provably identical to the target model alone.',
    what:['Draft phase: small 1B model generates gamma lookahead tokens with their full probability distributions','Verify phase: 7B target model evaluates all gamma draft tokens in a single parallel forward pass','Acceptance loop: P_accept equals min(1, P_target divided by P_draft) — rejection correction via residual sampling','Throughput benchmark: tokens per second at gamma = 2, 4, 8 versus sequential target-only baseline','Acceptance rate analysis: probability of acceptance at each lookahead position — expect decay at higher positions'],
    learn:['Why sequential decoding is memory-bandwidth-bound: weight reloading cost dominates per-token computation','Residual distribution sampling: how rejection correction maintains the exact target model output distribution','Optimal gamma selection: why more lookahead is not always better due to acceptance rate decay','Parallel verification batching: running the target model on gamma+1 positions in one single forward pass','Context truncation handling: position offset management when a cycle restarts at a rejection point'],
    weekplan:[
      {wk:'28',title:'Draft + Acceptance Logic',task:'Extract per-token logprob distributions from Ollama API. Draft phase: generate gamma tokens, store token IDs and probability distributions. Verify: single target model forward pass for gamma+1 positions. Acceptance loop: P_accept = min(1, P_target/P_draft), uniform sample, residual correction. Log acceptance rate per position across 10 example prompts.'},
      {wk:'29',title:'Benchmarking + Dashboard',task:'Baseline: measure tokens per second with target-only generation on 20 prompts at 200 tokens each. Speculative: run gamma = 2, 4, and 8 on identical prompts. Plot tokens per second vs gamma. Plot acceptance rate vs lookahead position. Validate output equivalence with same random seed. Identify breakeven rejection rate.'},
    ],
    hardPart:'Logprob extraction requires the logprobs API option in Ollama. Verify this is available in your Ollama version before starting Week 28. If unavailable, use the P6 PyTorch model directly — the acceptance math is identical.',
    tip:'The output distribution is provably identical to the target model alone — this is not an approximation. It is a mathematically exact result. Understanding why requires reading the residual distribution derivation carefully before implementing.',
    deliverable:'Benchmark dashboard showing 2x or greater throughput improvement with acceptance rate analysis report' },

  { id:'P11', weeks:'Wk 30–31', color:'9E6A3F', light:'FFF0E8', title:'Automated Evaluation (LLM-as-a-Judge) Framework', phase:'Phase IV  ·  Advanced Inference & MLOps',
    focus:'Pass@K Metric · Docker Sandboxed Execution · Chain-of-Thought Grading · Regression Dashboard',
    why:'Evaluating LLM quality by reading outputs is unscalable and inconsistent. This teaches you how production ML teams measure model quality — the same Pass@K metric used by OpenAI for HumanEval, and the same LLM-as-judge pattern used for RLHF reward model evaluation.',
    what:['Benchmark suite: 30 diverse tasks — 15 coding tasks with hidden unit tests and 15 reasoning tasks','Pass@K estimator: sample K solutions per task, execute in Docker sandbox, compute P(at least 1 passes)','Docker sandboxed execution: isolated Python container, no network access, mounted temp directory, execution timeout','LLM-as-judge pipeline: structured rubric (4 to 6 criteria, 1 to 5 scale), chain-of-thought format, JSON score output','Regression dashboard: PostgreSQL time series of scores — detect regressions when score drops 10% or more'],
    learn:['Pass@K formula: 1 minus C(n-c, K) divided by C(n, K) — the unbiased estimator for k-shot pass probability','Docker SDK execution: container lifecycle, temp directory mounting, stdout and stderr capture, timeout handling','Structured chain-of-thought grading: how rubric design affects judge model consistency and score calibration','Regression testing for AI systems: treating prompt and model changes as versioned software releases','Judge consistency measurement: inter-run score variance as a confidence interval for evaluation reliability'],
    weekplan:[
      {wk:'30',title:'Pass@K + Docker Sandbox',task:'Design 30-task benchmark suite: 15 coding tasks with hidden pytest unit tests, 15 reasoning tasks. Docker SDK: create container, mount temp directory, execute pytest with timeout, capture all results. Implement unbiased Pass@K estimator. Sample K=5 per coding task. Plot Pass@1 vs Pass@5 by task domain.'},
      {wk:'31',title:'LLM Judge + Dashboard',task:'Write 4-criteria evaluation rubrics for reasoning tasks as JSON schema. Judge pipeline: rubric prompt, model call, parse JSON score output. Validate consistency: run same task 3 times, score standard deviation below 0.5. PostgreSQL evaluation history schema. Regression dashboard. Run full 30-task suite and document results.'},
    ],
    hardPart:'Pre-pull the python:3.11-slim Docker image before testing. Cold image pulls take 15 to 30 seconds and distort execution time measurements. Always set a container timeout to prevent infinite-loop candidates from hanging the evaluation pipeline.',
    tip:'Run the P11 evaluation suite on the base model BEFORE P12 fine-tuning to establish a clean baseline. Without a pre-training benchmark, you cannot prove fine-tuning helped. P11 is the measurement instrument for the P12 experiment.',
    deliverable:'Evaluation framework with Pass@K pipeline, LLM-as-judge, and regression dashboard over 30-task benchmark suite' },

  { id:'P12', weeks:'Wk 32–34', color:'3F9E9E', light:'E8FFFF', title:'Knowledge Distillation & Fine-Tuning Pipeline', phase:'Phase IV  ·  Advanced Inference & MLOps',
    focus:'LoRA Low-Rank Adapters · PEFT · Adapter Merging · Synthetic Dataset Curation · 6GB VRAM Budget',
    why:'Full fine-tuning of a 7B model requires 112GB VRAM for weights, gradients, and optimiser states. LoRA reduces trainable parameters by 99% by injecting small A times B matrices at each linear layer while keeping base weights frozen. This is what allows customising foundation models on consumer hardware.',
    what:['Synthetic dataset: 500 or more instruction-response pairs collected from a teacher model API','LoRA configuration: rank r=8, alpha=16, target modules are q_proj and v_proj, trainable parameters under 1%','SFTTrainer: gradient accumulation, gradient checkpointing, paged AdamW — all staying within 6GB VRAM','Adapter merging: W_merged = W plus B times A times (alpha/r) — zero inference overhead after merge','P11 evaluation: measure Pass@K and judge score improvement over untuned base model baseline'],
    learn:['LoRA mathematics: why B times A with small r approximates the full weight update at 1% of parameter cost','Frozen base weights: gradient flow stops at frozen parameters, only A and B matrices receive gradient updates','Knowledge distillation: why teacher-generated synthetic data outperforms raw web data for narrow task domains','VRAM budget management: gradient checkpointing, paged AdamW, and mixed-precision training as orthogonal strategies','Adapter merging arithmetic: W_merged = W plus B times A times (alpha/r) and why this normalises the contribution'],
    weekplan:[
      {wk:'32',title:'Dataset + LoRA Setup',task:'Design 20 domain-specific task templates. Query teacher model API for 500 or more instruction-response pairs. Filter dataset: remove duplicates, length outliers, and incoherent responses. Apply LoraConfig with r=8, alpha=16, target_modules=[q_proj, v_proj]. Verify trainable parameter count is under 1% of total model parameters.'},
      {wk:'33',title:'Training + Adapter Merge',task:'SFTTrainer: batch_size=4, gradient_accumulation=8, gradient_checkpointing=True. Train for 3 epochs. Plot training and validation loss curves. Verify convergence and check for overfitting. Generate output samples at checkpoints. Merge adapters: W_merged = W plus B times A times (alpha/r). Verify merged model parameter count matches base model exactly.'},
      {wk:'34',title:'Evaluation + Ablation',task:'Run P11 suite on merged fine-tuned model: Pass@K for coding tasks, judge score for reasoning tasks. Compute improvement delta vs untuned baseline. Ablation: r=4 vs r=8 vs r=16 — plot accuracy vs trainable parameter count tradeoff. Write technical report. Package pipeline as reusable training script.'},
    ],
    hardPart:'Dataset quality dominates outcome quality above all other factors. 200 high-quality curated examples outperform 2,000 noisy ones every time. Invest significant time in Week 32 curation — deduplicate aggressively and manually review 50 random samples before training begins.',
    tip:'Run P11 on the base model BEFORE fine-tuning. Without a pre-training baseline, you cannot prove fine-tuning helped. The evaluation framework was built precisely to measure this improvement delta.',
    deliverable:'LoRA-tuned model with measurable Pass@K improvement, training loss curves, and reusable training pipeline script' },

  { id:'P13', weeks:'Wk 35–37', color:'7A3F9E', light:'F0E8FF', title:'GPU SRAM Flash-Attention Simulator', phase:'Phase V  ·  Hardware & Memory Optimisation',
    focus:'HBM vs SRAM IO Analysis · Tiling Architecture · Online Softmax · O(N squared) to O(N) IO Reduction',
    why:'Standard attention materialises an N times N score matrix in slow GPU global memory. At sequence length 4096, this is 128MB per attention head. Flash-Attention eliminates this materialisation by restructuring computation into SRAM-resident tiles. Understanding this separates hardware-aware engineers from framework users.',
    what:['Standard attention with IO instrumentation: count every HBM read and write event, verify O(N squared) growth','Tiling algorithm: divide Q, K, and V into blocks sized to fit within the SRAM budget, process block pairs','Online softmax: running m_i maximum and l_i normaliser updated per K/V tile without materialising the full row','Rescaling mechanism: when a new maximum is found, previous running output is rescaled by exp(m_old minus m_new)','IO comparison chart: HBM accesses vs sequence length for standard attention versus Flash-Attention tiling'],
    learn:['GPU memory hierarchy: SRAM is on-chip and fast at approximately 192KB; HBM is off-chip and slow at 80GB bandwidth','Why large softmax matrices cause IO bottlenecks: every element must travel across the slow memory bus','Online log-sum-exp trick: numerically stable softmax computation without materialising the full score row','IO complexity reduction: from O(N squared) to O(N) by restructuring computation order, not changing the math','Memory-bound vs compute-bound: why modern transformers are IO-limited, not FLOP-limited, on real GPU hardware'],
    weekplan:[
      {wk:'35',title:'Standard Attention + IO Tracking',task:'Standard scaled dot-product attention in NumPy. IO counter increments on every array read and write. Run at sequence lengths N = 128, 256, 512, 1024, 2048. Plot IO count vs N. Verify O(N squared) growth curve. Identify score matrix materialisation as the dominant IO contributor.'},
      {wk:'36',title:'Flash-Attention Tiling',task:'Read Flash-Attention paper (Dao et al., 2022) Sections 3 and 4 BEFORE writing any code. Implement: Q block outer loop, K and V block inner loop. Online softmax with running m_i, l_i, and output rescaling. Validate: output matches standard attention output within 1e-5 absolute tolerance. Run at same N values. Plot Flash-Attention IO on same chart.'},
      {wk:'37',title:'Tile Analysis + Dashboard',task:'Vary tile size (32, 64, 128, 256 elements). Implement SRAM budget constraint parameter. Chart IO savings percentage across N and tile size combinations. Build 4-panel dashboard: IO comparison, savings percentage, tile sensitivity, memory lifecycle. Write analysis: optimal tile size for your SRAM budget and breakeven sequence length.'},
    ],
    hardPart:'Read the Flash-Attention paper (Dao et al., 2022) Sections 3 and 4 BEFORE writing any code in Week 36. The online softmax m_i and l_i update with rescaling is non-obvious and produces incorrect results if implemented from intuition alone.',
    tip:"Flash-Attention's key insight: correct softmax output is computable without ever seeing all scores simultaneously. The running max and normaliser allow correcting past outputs as new maximums arrive. Build intuition for online softmax before touching the tiling loop.",
    deliverable:'Flash-Attention simulator with IO instrumentation, O(N) vs O(N squared) comparison chart, and 4-panel dashboard' },

  { id:'P14', weeks:'Wk 38–40', color:'9E3F4A', light:'FFE8EA', title:'Heterogeneous Memory Layer & Model Offloader', phase:'Phase V  ·  Hardware & Memory Optimisation',
    focus:'PyTorch Forward Hooks · PCIe Async Pre-Fetching · CPU/VRAM Layer Sharding · VRAM Floor Cap',
    why:'A 13B parameter model at FP16 requires 26GB VRAM. Consumer GPUs have 6 to 12GB. Heterogeneous offloading is how llama.cpp and Hugging Face accelerate run models larger than VRAM by sharding across CPU RAM and GPU — layer by layer, with async pre-fetching to hide PCIe transfer latency.',
    what:['Layer partitioner: inspect model layer sizes, assign each to CUDA or CPU based on VRAM_BUDGET_GB parameter','pre_forward_hook: move CPU-resident layer to GPU immediately before its forward pass executes','post_forward_hook: return layer to CPU immediately after its forward pass completes','Async pre-fetcher: background thread starts CUDA stream transfer of layer i+1 while layer i is executing','Benchmark: synchronous vs async offloading tokens per second plus VRAM Floor Cap measurement'],
    learn:['PyTorch hook system: pre_forward_hook and post_forward_hook as non-intrusive module interceptors','PCIe bus bandwidth: approximately 16 GB/s for PCIe 4.0 x16 and how it governs offloading feasibility','CUDA streams: non-blocking transfers that overlap with active compute to hide PCIe transfer latency','VRAM Floor Cap: minimum VRAM equals the two largest consecutive layers (current plus pre-fetched next)','Hook attachment scope: attach only to leaf modules such as Linear and LayerNorm, never to container modules'],
    weekplan:[
      {wk:'38',title:'Partitioner + Hooks',task:'Inspect model: log layer-by-layer parameter count and memory footprint in MB. Build partitioner: assign layers to cuda or cpu based on VRAM_BUDGET_GB. Implement pre_forward_hook: call layer.cuda() before forward. Implement post_forward_hook: call layer.cpu() after. Validate: offloaded output matches full GPU-resident output via torch.allclose. Measure peak VRAM with nvidia-smi.'},
      {wk:'39',title:'Synchronous Benchmark',task:'Measure per-layer CPU to GPU transfer latency using CUDA Event timing. Plot transfer time vs parameter count — estimate effective PCIe bandwidth. Run full inference with synchronous offloading: measure tokens per second. Identify the slowest bottleneck layers. Experiment: VRAM budgets of 2GB, 4GB, 6GB — measure tokens per second at each level.'},
      {wk:'40',title:'Async Pre-Fetcher + Dashboard',task:'Pre-fetch thread: CUDA stream transfer of layer i+1 during layer i execution. Synchronisation barrier in pre_forward_hook: wait for stream completion before forward pass begins. Benchmark async vs synchronous tokens per second. Track VRAM Floor Cap across configurations. Build 4-panel dashboard: VRAM vs config, tokens per second vs config, PCIe latency, layer execution timeline.'},
    ],
    hardPart:'Hooks on container modules such as Sequential and ModuleList fire for every child module inside them, causing double-transfers. Attach hooks exclusively to leaf modules — individual Linear, LayerNorm, and Embedding layers only.',
    tip:'This project connects everything: KV Cache (P8), quantization (P8), Flash-Attention (P13), and layer offloading (P14) all address the same hardware memory constraint from different angles. Document these connections explicitly in your README — the synthesis is the proof of mastery.',
    deliverable:'Async offloading engine with VRAM Floor Cap analysis, synchronous vs async benchmark, and 4-panel dashboard' },

  { id:'P15', weeks:'Wk 41–43', color:'2E6E8E', light:'E0F4FF', title:'High-Throughput Timeseries Chunk Compressor', phase:'Phase VI  ·  Byte-Level Systems',
    focus:'Zstd Dictionary Training · Lossless Streaming · Error-Bounded Lossy Quantization · Binary Serialization',
    why:'Compression on isolated small packets is inefficient because the compressor has no context to exploit repeating patterns. Dictionary pre-loading solves this by training a shared schema fingerprint from representative samples — the same technique used by Cloudflare and Meta for high-frequency telemetry pipelines. Building it from scratch teaches you the exact tradeoffs between compression ratio, CPU cost, and latency at byte level.',
    what:[
      'Analytical dictionary synthesizer: compile repeating JSON schema samples into a static binary Zstd dictionary file',
      'Error-bounded lossy pre-processor: filter floating-point metric arrays via linear delta prediction, capping reconstruction error at 0.01%',
      'Streaming binary serialization layer: sliding-window compression loop feeding compressed bytes continuously into network buffers',
      'Performance benchmark: chart compression ratio vs CPU processing time across packet sizes and dictionary sizes',
    ],
    learn:[
      'Why compression fails on isolated small packets: the cold-start problem and how dictionary pre-loading provides shared context',
      'Zstd dictionary training algorithm: how representative sample batches are distilled into a reusable binary index structure',
      'Error-bounded lossy quantization: delta prediction, bit truncation, and guaranteeing maximum reconstruction error bounds',
      'Storage volume vs CPU execution time tradeoff: measuring the compression-compute frontier for real-time streaming workloads',
      'Binary serialization design: sliding-window buffer management, memory alignment, and zero-copy streaming patterns',
      'libzstd C API via Python ctypes: ZSTD_CCtx, ZSTD_DCtx, dictionary loading, and streaming frame management',
    ],
    weekplan:[
      {wk:'41',title:'Static Dictionary Synthesis',
       task:'Gather 10,000 representative metric JSON samples. Invoke Zstd dictionary training via libzstd (ZDICT_trainFromBuffer). Measure dictionary size vs compression gain across varying sample counts (100, 1K, 10K). Benchmark: compressed packet size with vs without dictionary at identical payloads. Plot optimization level (1 to 22) vs speed tradeoff curve.'},
      {wk:'42',title:'Error-Bounded Lossy Quantization',
       task:'Implement delta prediction filter on float32 time-series arrays: compute differences between consecutive values, round trailing bits, store residuals. Enforce 0.01% max reconstruction error: validate by re-expanding and comparing against original values. Measure compression ratio uplift vs lossless baseline. Build unit tests asserting error bounds pass for 1,000 synthetic metric streams.'},
      {wk:'43',title:'Streaming Buffer Serialization',
       task:'Construct sliding-window compression loop: chunk incoming bytes into fixed blocks, compress each block with pre-loaded dictionary, write to output buffer. Implement ZSTD streaming API (ZSTD_compressStream2) for continuous frame generation. Plot throughput in MB/s vs window size. Benchmark end-to-end: raw JSON bytes in, compressed binary stream out, decompressed and verified on the other end.'},
    ],
    hardPart:'Dictionary quality degrades sharply if training samples are not representative of production data. Use actual metric payloads from your RunTrack or GastroHub data schemas as training input — not synthetic random data. A dictionary trained on wrong patterns can make compression worse than no dictionary at all.',
    tip:'This project feeds directly into both Capstone projects. Capstone 02 (Industrial Copilot) uses P15 to compress IoT telemetry streams. Capstone 01 (Hydra Engine) uses P15 Zstd dictionaries to compress KV-Cache transfers over the PCIe bus from P14. Build P15 with a clean compression/decompression API so both capstones can import it as a library module.',
    deliverable:'Zstd dictionary compressor library + error-bounded lossy quantizer + streaming benchmark showing throughput in MB/s' },

  { id:'P16', weeks:'Wk 44–46', color:'3E6B3E', light:'E8F5E8', title:'Distributed High-Availability Job Scheduler', phase:'Phase VI  ·  Distributed Coordination',
    focus:'Optimistic Row-Level Locking · Redis ZSET Indexes · Lease-Based Worker Ownership · Heartbeat Trackers',
    why:'At scale, cron jobs and naive task queues fail catastrophically: two workers fire the same job simultaneously, crashed workers leave tasks orphaned forever, and hot database polling saturates the primary under load. This project builds a production-grade distributed scheduler — the same architecture inside Airflow, Temporal, and GitHub Actions — from first principles, teaching you every failure mode and its precise mitigation.',
    what:[
      'Multi-node scheduler: poll a shared PostgreSQL jobs table indexed via Redis ZSETs for O(log N) time-indexed lookups',
      'Lease-management worker fleet: atomic compare-and-swap (CAS) row locks via UPDATE WHERE status=SCHEDULED prevent duplicate execution',
      'Heartbeat audit monitor: background thread detects crashed workers and automatically releases and reschedules orphaned tasks',
      'Dead-letter recovery pipeline: isolate permanently failing jobs into a separate queue with full error context for manual inspection',
      'Load simulation: fire 1 million scheduled jobs across 10 worker nodes, assert zero duplicate executions and zero permanent orphans',
    ],
    learn:[
      'The scheduling paradox: the fundamental tradeoff between at-least-once delivery (durability) and exactly-once execution (no double-firing)',
      'Optimistic concurrency control: why UPDATE WHERE status=SCHEDULED is safer than SELECT FOR UPDATE under high concurrency',
      'Redis ZSET as a time index: ZADD with Unix timestamp score, ZRANGEBYSCORE for O(log N) due-job queries at any time horizon',
      'Lease-based worker ownership: atomic acquisition, TTL expiry, and renewal mechanics for distributed task claiming',
      'Heartbeat monitoring: detecting worker liveness, calculating stale lease thresholds, and safe re-enqueue without data loss',
      'Dead-letter queue design: separating permanently failed jobs from transient failures to prevent retry storms',
    ],
    weekplan:[
      {wk:'44',title:'Time-Indexed Sorted Store',
       task:'Docker Compose: PostgreSQL jobs table (id, payload, status, scheduled_at, worker_id, lease_expires_at, retry_count) plus Redis. Seed 100,000 scheduled jobs with varying future timestamps. Implement Redis ZSET index: ZADD jobs:schedule {unix_ts} {job_id} on insert. Build scheduler poll loop: ZRANGEBYSCORE to find due jobs, O(log N) per query. Benchmark: ZSET poll latency at 100K, 1M, 10M entries.'},
      {wk:'45',title:'Atomic Lease Acquisition Loop',
       task:'Implement CAS lock: UPDATE jobs SET status=PROCESSING, worker_id={id}, lease_expires_at=NOW()+30s WHERE id={job_id} AND status=SCHEDULED. Check rows_affected == 1 to confirm exclusive acquisition — if 0, another worker claimed it. Build worker pool: 10 concurrent workers each running the CAS loop. Run 50,000 concurrent job claims — assert every job claimed exactly once. Log contention rate (CAS failures per second).'},
      {wk:'46',title:'Zombie Process Extermination',
       task:'Implement heartbeat: each worker sends UPDATE jobs SET lease_expires_at=NOW()+30s WHERE worker_id={id} AND status=PROCESSING every 10 seconds. Build heartbeat monitor: background thread queries SELECT * FROM jobs WHERE status=PROCESSING AND lease_expires_at < NOW(). For stale leases: UPDATE status=SCHEDULED, worker_id=NULL (returns job to queue). Test zombie scenario: kill a worker mid-job, verify monitor detects and re-queues within 35 seconds. Build dead-letter queue: jobs exceeding retry_count=5 routed to failed_jobs table.'},
    ],
    hardPart:'The most dangerous edge case is a worker that is slow, not crashed — it is still alive but its lease has expired and another worker has claimed the job. Now two workers are executing the same job. Defense: make all job handlers idempotent (safe to run twice) before relying solely on lease protection. Idempotency is your last line of defence.',
    tip:'This project integrates directly into Capstone 01 (Hydra Engine) as the job management layer for routing inference requests across Prefill and Decode nodes, and into Capstone 02 (Industrial Copilot) for scheduling factory inspection sequences. Design the scheduler as a standalone importable module from Week 44 — both capstones will depend on it.',
    deliverable:'Distributed scheduler with CAS lease acquisition, heartbeat zombie detection, dead-letter queue, and 1M-job load test' },

  { id:'P17', weeks:'Wk 47–52', color:'5E3E7A', light:'F0E8FF', title:'Distributed Multi-GPU Tensor Parallelism Engine', phase:'Phase VII  ·  Advanced Deep Learning',
    focus:'Column vs Row Parallelism · NCCL Collective Kernels · All-Reduce Synchronization · torch.distributed',
    why:'Training and serving billion-parameter models on a single GPU is impossible — the weights alone exceed available VRAM. Tensor parallelism is how Megatron-LM, DeepSpeed, and vLLM shard individual matrix operations across multiple devices. Understanding column and row parallelism, All-Reduce communication patterns, and interconnect bandwidth costs is what separates engineers who can scale AI systems from those who can only run them.',
    what:[
      'Column parallelism implementation: slice weight matrices W_gate and W_up vertically, route each shard to an independent worker process',
      'Row parallelism implementation: shard down-projection W_down horizontally across devices, accumulate partial outputs',
      'All-Reduce kernel: multi-process loop aggregating and redistributing partial tensor states across all nodes using raw collective communications',
      'torch.distributed integration: configure multi-device process groups, profile communication overhead vs compute time',
      'Interconnect bandwidth profiling: measure All-Reduce latency across 2, 4, and 8 simulated nodes, identify the communication bottleneck point',
    ],
    learn:[
      'Column parallelism math: slicing W vertically means each device computes Y_i = X * W_i — outputs are concatenated, not summed',
      'Row parallelism math: slicing W horizontally means each device computes Y_i = X_i * W — outputs require All-Reduce summation',
      'All-Reduce collective: the ring-AllReduce algorithm, why it scales to O(1) bandwidth per node regardless of node count',
      'All-Gather vs All-Reduce: when each collective is appropriate in the forward vs backward pass of tensor-parallel layers',
      'Tensor parallelism vs pipeline parallelism: intra-layer splits (TP) vs inter-layer splits (PP) and when to combine them',
      'NCCL communication overhead: measuring the gap between compute time and synchronization wait time as a function of tensor size',
    ],
    weekplan:[
      {wk:'47',title:'Column Parallelism Sharding',
       task:'Set up torch.distributed process group with 4 simulated workers (torchrun --nproc_per_node=4). Extract linear layer weight matrices from P6 transformer. Implement column split: torch.chunk(W, num_devices, dim=1). Each process receives its W_i shard and computes local Y_i = X * W_i. Verify: concatenating all Y_i across workers produces identical output to single-device computation. Benchmark: 2 vs 4 workers on matrix sizes 4096x4096 and 8192x8192.'},
      {wk:'48',title:'Row Parallelism + Partial Accumulation',
       task:'Implement row split: torch.chunk(W, num_devices, dim=0). Split input X accordingly: each device receives X_i shard. Compute partial output Y_i = X_i * W_i on each device. Implement All-Reduce step: dist.all_reduce(Y, op=dist.ReduceOp.SUM) to accumulate partial outputs into final Y. Validate numerical equivalence against single-device reference. Combine column and row parallelism: implement a complete feedforward block sharded across 4 devices.'},
      {wk:'49',title:'All-Reduce Kernel Assembly',
       task:'Implement ring-All-Reduce manually using point-to-point dist.send and dist.recv calls — do not use dist.all_reduce. Divide tensor into num_devices chunks. Run reduce-scatter phase: each node sends its chunk to the right, accumulates received chunk. Run all-gather phase: distribute the fully-reduced chunk. Validate: manual ring-AllReduce produces same result as dist.all_reduce. Measure latency: manual vs NCCL-backed dist.all_reduce for tensor sizes 1MB, 10MB, 100MB.'},
      {wk:'50',title:'Interconnect Profiling + Dashboard',
       task:'Configure multi-device tests using torch.distributed with 2, 4, and 8 process groups. Profile with torch.profiler: separate compute time vs All-Reduce communication time per forward pass. Plot: communication overhead percentage vs number of devices. Identify the tensor size threshold where communication dominates compute. Build 4-panel dashboard: latency vs device count, bandwidth utilisation, compute vs comms breakdown, scaling efficiency (ideal vs actual speedup).'},
    ],
    hardPart:'NCCL requires all processes in a communication group to call collective operations in the same order and at the same time — a single process that skips an All-Reduce causes the entire group to deadlock indefinitely. Always wrap collective calls in a try/except and ensure every code path through your model reaches the same set of collective calls regardless of conditional branches.',
    tip:'If you do not have access to multiple physical GPUs, simulate multi-device tensor parallelism using CPU process groups with gloo backend (dist.init_process_group("gloo")). The communication patterns, ring-AllReduce math, and profiling methodology are completely identical — only the bandwidth numbers differ. Run the real GPU benchmark when available, but do not block the project on hardware access.',
    deliverable:'Tensor parallelism engine with column + row sharding, manual ring-AllReduce, and 4-panel interconnect profiling dashboard' },
];

// Build children
const children = [];

// Cover
children.push(
  sp(1440),
  p('ARVEEND  ·  SELF-LEARN ENGINEERING PROGRAMME  ·  2025–2026', {bold:true,color:'6B7EA8',size:18,align:AlignmentType.CENTER,after:200}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:120},children:[new TextRun({text:'Systems & AI Engineering',bold:true,color:'1B2A4A',size:56,font:'Arial'})]}),
  new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:120},children:[new TextRun({text:'Master Roadmap',bold:true,color:'2E5090',size:44,font:'Arial'})],border:{bottom:{style:BorderStyle.SINGLE,size:10,color:'2E5090',space:8}}}),
  sp(400),
  tbl([row([
    cell([p('14',{bold:true,color:'FFFFFF',size:52,align:AlignmentType.CENTER,after:16}),p('Projects',{color:'C0D0F0',size:20,align:AlignmentType.CENTER,after:0})],{fill:'1B2A4A',w:3120,bc:'1B2A4A'}),
    cell([p('40',{bold:true,color:'FFFFFF',size:52,align:AlignmentType.CENTER,after:16}),p('Weeks',{color:'C0D0F0',size:20,align:AlignmentType.CENTER,after:0})],{fill:'2E5090',w:3120,bc:'2E5090'}),
    cell([p('5',{bold:true,color:'FFFFFF',size:52,align:AlignmentType.CENTER,after:16}),p('Phases',{color:'C0D0F0',size:20,align:AlignmentType.CENTER,after:0})],{fill:'3B6FBF',w:3120,bc:'3B6FBF'}),
  ])],[3120,3120,3120]),
  sp(400),
  p('From backend infrastructure and AI orchestration to raw transformer mathematics and hardware-level memory optimisation.', {size:21,color:'5A6A8A',italic:true,align:AlignmentType.CENTER}),
  pb()
);

// Intro section
children.push(
  h1('How to Use This Roadmap'),
  p('This document is your full execution guide for the 14-project, 40-week programme. Each project section covers: why you are building it, exactly what you will build, what you will learn, a week-by-week breakdown, the hardest implementation challenge, and the deliverable that proves completion.', {size:21,color:'333355',after:160}),
  h2('Programme Structure at a Glance'),
  tbl([
    row([
      cell([p('Phase',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:700,bc:'1B2A4A'}),
      cell([p('Title',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:2500,bc:'1B2A4A'}),
      cell([p('Projects',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:1100,bc:'1B2A4A'}),
      cell([p('Weeks',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:1000,bc:'1B2A4A'}),
      cell([p('Core Skills',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:4060,bc:'1B2A4A'}),
    ]),
    ...[
      ['I','High-Performance Backend Infrastructure','P1–P4','Wk 1–12','Vector databases, distributed caching, atomic concurrency, WebSocket real-time systems'],
      ['II','Agentic Orchestration & AI Workspaces','P5','Wk 13–15','Multi-LLM orchestration, BullMQ dependency chains, human-in-the-loop governance'],
      ['III','Core Neural Systems Engineering','P6–P8','Wk 16–25','Transformer internals from tensors, ANN indexing algorithms, model quantization and KV caching'],
      ['IV','Advanced Inference Scale & MLOps','P9–P12','Wk 26–34','MoE sparse routing, speculative decoding, automated evaluation, LoRA fine-tuning'],
      ['V','Hardware & Memory Optimisation','P13–P14','Wk 35–40','Flash-Attention IO tiling, PCIe layer offloading, CPU/VRAM heterogeneous execution'],
    ].map(([ph,name,projs,wks,skills],i) => row([
      cell([p(ph,{bold:true,color:'2E5090',size:20,after:0,align:AlignmentType.CENTER})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:700,bc:'CCCCCC'}),
      cell([p(name,{bold:true,size:19,color:'1A1A2E',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:2500,bc:'CCCCCC'}),
      cell([p(projs,{size:19,color:'3B6FBF',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:1100,bc:'CCCCCC'}),
      cell([p(wks,{size:19,color:'5A6A8A',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:1000,bc:'CCCCCC'}),
      cell([p(skills,{size:18,color:'444466',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:4060,bc:'CCCCCC'}),
    ]))
  ],[700,2500,1100,1000,4060]),
  sp(160),
  h2('Three Rules for the Entire Programme'),
  bullet('Mock everything external on Day 1','Every project depends on something external (Finnhub, Ollama APIs, Docker). Mirror all responses to local JSON fixtures the first time they are called. This prevents rate limits and network issues from blocking momentum across the entire 40 weeks.'),
  bullet('Build the broken version first','P3 teaches you to build the race-condition-prone checkout before the fix. Apply this everywhere: understand the problem empirically before solving it. A working before state makes the solution meaningful and memorable.'),
  bullet('Deliver something visible every week','Each week must produce a running endpoint, a chart, a passing test, or a Loom recording. Invisible progress accumulates technical debt. Visible progress builds momentum and portfolio evidence simultaneously.'),
  sp(120),
  h2('How the 14 Projects Connect to Each Other'),
  bullet('','P4 (Arveend Terminal) integrates P1 (RAG for earnings summaries), P2 (Cache-Aside for stock profiles), and P3 (atomic Lua for paper trading). Build P1 through P3 knowing they will wire together in P4.'),
  bullet('','P9 (MoE Transformer) directly extends P6 (Transformer from Scratch). The MoE layer is a drop-in replacement for the feedforward block. P6 must be working before P9 begins.'),
  bullet('','P11 (Evaluation Framework) is the measurement instrument for P12 (Fine-Tuning). Run P11 on the base model before P12 fine-tuning, then again after — the delta is your proof of improvement.'),
  bullet('','P13 (Flash-Attention) and P14 (Offloading) both address GPU memory limits from different directions. Together they form a complete hardware optimisation toolkit — document the connection explicitly.'),
  sp(240),
  pb()
);

// Prep
children.push(
  h1('Preparation Checklist: Complete Before Week 1'),
  p('Complete every item below before writing any project code. Environment issues discovered in Week 3 cost more time than setting everything up correctly in Week 0.', {size:21,color:'333355',after:160}),
  tbl([
    row([
      cell([p('Setup Item',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:2600,bc:'1B2A4A'}),
      cell([p('What to Do Exactly',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:4200,bc:'1B2A4A'}),
      cell([p('First Used In',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:2560,bc:'1B2A4A'}),
    ]),
    ...[
      ['Docker Desktop','Install Docker Desktop. Verify with: docker run --rm postgres:16 echo ok','P1 Week 2'],
      ['Ollama + Base Models','Install Ollama. Pull: llama3.2:3b, qwen2.5-coder:7b, nomic-embed-text','P1 Week 1'],
      ['Phase II Agent Models','Pre-pull deepseek-r1:7b and llama3.1:8b (large downloads, start early)','P5 Week 13'],
      ['GitHub Monorepo','Create repo with folders: /p1-rag, /p2-cache, /p3-mutex, /p4-terminal, etc.','P1 Week 1'],
      ['Finnhub API Key','Register free account at finnhub.io. Store key in root .env file.','P4 Week 1'],
      ['k6 Load Testing','Install k6 v0.51+. Verify: k6 run --help returns usage output.','P2 Week 8'],
      ['PyTorch and NumPy','pip install torch numpy. Verify with python -c "import torch; print(torch.__version__)"','P6 Week 16'],
      ['Hugging Face PEFT','pip install peft transformers trl accelerate bitsandbytes (for P12)','P12 Week 32'],
      ['Obsidian or Notion Log','Weekly template: Week N | Built | Broke | Fixed | Goal for next week','Every week'],
      ['Mock Fixture Directory','Create /mocks folder in every project directory before writing any API call code','P4 Week 1'],
    ].map(([item,what,first],i) => row([
      cell([p(item,{bold:true,size:19,color:'2E5090',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:2600,bc:'CCCCCC'}),
      cell([p(what,{size:19,color:'333355',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:4200,bc:'CCCCCC'}),
      cell([p(first,{size:18,color:'6B7EA8',italic:true,after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:2560,bc:'CCCCCC'}),
    ]))
  ],[2600,4200,2560]),
  sp(240),
  pb()
);

// Projects
for (const proj of projects) {
  children.push(
    tbl([row([
      cell([p(proj.id,{bold:true,color:'FFFFFF',size:32,align:AlignmentType.CENTER,after:10}),p(proj.weeks,{color:'D0DCF0',size:18,align:AlignmentType.CENTER,after:0})],{fill:proj.color,w:1440,bc:proj.color}),
      cell([p(proj.title,{bold:true,color:'FFFFFF',size:26,after:24}),p(proj.phase,{color:'A0C0E8',size:18,italic:true,after:0})],{fill:'1B2A4A',w:CW-1440,bc:'1B2A4A'}),
    ])],[1440,CW-1440]),
    sp(60),
    p(proj.focus,{size:20,color:'5A6A8A',italic:true,after:160}),
    h2('Why This Project Exists', proj.color),
    p(proj.why,{size:21,color:'2A2A44',after:160}),
    h2('What You Are Building', proj.color),
  );
  for (const w of proj.what) children.push(bullet('', w));
  children.push(sp(120), h2('What You Will Learn', proj.color));
  for (const l of proj.learn) children.push(num(l));
  children.push(sp(120), h2('Week-by-Week Execution Plan', proj.color));
  children.push(tbl([
    row([
      cell([p('Wk',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:proj.color,w:800,bc:proj.color}),
      cell([p('Focus Area',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:proj.color,w:1900,bc:proj.color}),
      cell([p('Concrete Tasks',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:proj.color,w:CW-2700,bc:proj.color}),
    ]),
    ...(proj.weekplan||[]).map((wk,i)=>row([
      cell([p(wk.wk,{bold:true,color:proj.color,size:20,after:0,align:AlignmentType.CENTER})],{fill:i%2===0?proj.light:'FFFFFF',w:800,bc:'CCCCCC'}),
      cell([p(wk.title,{bold:true,size:19,color:'1A1A2E',after:0})],{fill:i%2===0?proj.light:'FFFFFF',w:1900,bc:'CCCCCC'}),
      cell([p(wk.task,{size:19,color:'333355',after:0})],{fill:i%2===0?proj.light:'FFFFFF',w:CW-2700,bc:'CCCCCC'}),
    ]))
  ],[800,1900,CW-2700]));
  children.push(sp(160), h2('Critical Implementation Notes', proj.color));
  children.push(noteBox('CAUTION', proj.hardPart, '7A5500','5A3A00','FFF8E0','E0C050'));
  children.push(sp(60));
  children.push(noteBox('TIP', proj.tip, '1A6A3A','0A4A2A','F0FFF8','50C080'));
  children.push(sp(120), h2('Deliverable: Proof of Completion', proj.color));
  children.push(tbl([row([cell([p('Deliverable:  '+proj.deliverable,{bold:true,size:20,color:'FFFFFF',after:0})],{fill:proj.color,w:CW,bc:proj.color})])],[CW]));
  children.push(sp(240), pb());
}

// ── Capstone sections ────────────────────────────────────────────
const capstones = [
  {
    id:'CAPSTONE 01', badge:'C1', duration:'Duration: 1 Year', color:'1A3A6A', light:'E0EAFF',
    title:'The Hydra Engine: Disaggregated Multi-Tenant Inference Cluster',
    subtitle:'Phase Splitting (Prefill vs Decode) · Distributed PagedAttention KV-Pool · Continuous Batching · NIXL Transport',
    what:[
      'Enterprise-grade serving platform: separates compute-bound Prefill clusters from memory-bound Decode nodes over low-latency networks',
      'Continuous Batching controller pipeline: automated request scheduling that maximises GPU utilisation by filling decode slots dynamically',
      'Distributed KV-Cache streaming engine: powered by NIXL connectors and RDMA memory transfers between Prefill and Decode tiers',
      'Multi-tenant concurrency barriers: combines P3 atomic mutex handlers with P16 lease-based job scheduler for request routing',
    ],
    integrations:[
      {label:'Frontend Gateway', desc:'P4 Terminal and P5 Agent canvas channels monitor live text generation streams via WebSocket connections'},
      {label:'Prefill Compute Layer', desc:'Prefill arrays parse documents using P1 pgvector structures, running inputs through P13 Flash-Attention tiling blocks'},
      {label:'Decode Execution Pool', desc:'Decode nodes manage state pools using P8 INT4 quantized weight configurations with P10 speculative lookahead passes'},
      {label:'Data Transport Layer', desc:'KV-Cache allocations are serialized, compressed via P15 Zstd dictionaries, and streamed across P14 PCIe bus pathways'},
    ],
    concepts:[
      'Prefill vs Decode disaggregation: why compute-bound and memory-bound phases have fundamentally different hardware requirements',
      'PagedAttention KV-Pool: non-contiguous virtual memory for KV-Cache blocks, eliminating fragmentation in multi-tenant serving',
      'Continuous Batching: how in-flight request insertion maximises GPU utilisation versus static batch execution',
      'NIXL transport connectors: high-throughput inter-node tensor movement using RDMA-capable network interfaces',
      'Multi-tenant isolation: request routing, priority queuing, and resource quota enforcement across concurrent users',
    ],
    stack:'Go / C++ Core · vLLM Architecture · NIXL Tensor Transport · Continuous Batching Scheduler · PagedAttention Pool',
    hardPart:'The Prefill-Decode split introduces a KV-Cache migration problem: the KV tensors computed during Prefill must be transferred to the correct Decode node before generation begins. This transfer must complete before the first output token is emitted, making it latency-critical. Design the NIXL transfer pipeline to overlap with the tail of the Prefill computation.',
    tip:'Treat this capstone as an integration test for the entire programme. Every sub-system you built (P1 through P16) has a specific socket in this architecture. Map each integration point explicitly in your architecture diagram before writing a single line of capstone code. The diagram is the design.',
  },
  {
    id:'CAPSTONE 02', badge:'C2', duration:'Duration: 1 Year', color:'2A4A1A', light:'E8F5E0',
    title:'Industrial AI Operations Copilot for Manufacturing',
    subtitle:'Edge Computer Vision · Industrial IoT MQTT Streaming · Timeseries Anomaly Detection · Fault-Tolerant Local Execution',
    what:[
      'Local edge automation infrastructure: tracks multi-facility plant metrics and live video streams in real-time without cloud dependency',
      'Low-latency MQTT telemetry handlers: processes high-frequency physical equipment signals from IoT controllers over MQTT broker clusters',
      'Machine vision defect detection: local computer vision models monitor camera feeds for defect thresholds and asset tracking vectors',
      'Automated anomaly alerting circuit: routes anomaly warnings down to localised multi-channel dispatch frameworks instantly on detection',
    ],
    integrations:[
      {label:'Hardware Ingestion Layer', desc:'IoT controllers feed sensory metrics directly into the P15 Zstd chunk compressor for efficient stream buffering and transmission'},
      {label:'Distributed Scheduling', desc:'P16 lease-based job scheduler distributes factory machine inspection sequences and manages log rotation across all edge nodes'},
      {label:'Localised AI Analytics', desc:'P8 INT4 quantized vision model weights run computer-vision classifications entirely on-device without cloud round-trips'},
      {label:'Fault-Tolerant Cache Layer', desc:'P2 Cache-Aside and P3 atomic mutex patterns preserve factory log data and sensor state during sudden site network outages'},
    ],
    concepts:[
      'Edge inference deployment: running quantized vision models on constrained hardware without GPU availability',
      'MQTT broker architecture: topic hierarchy design, QoS levels, and broker clustering for high-frequency industrial telemetry',
      'Timeseries anomaly detection: statistical threshold methods, sliding-window Z-score, and isolation forest for equipment fault detection',
      'Fault-tolerant local execution: offline-first design patterns, local queue buffering, and sync-on-reconnect for unreliable site networks',
      'Multi-facility coordination: centralised dashboard aggregating edge node telemetry from geographically distributed plant sites',
    ],
    stack:'Python / C++ Core · MQTT Broker Clusters · Edge Computer Vision · Timeseries Anomaly Engine · Industrial IoT Interfaces',
    hardPart:'Industrial network environments are hostile to assumptions made during development: intermittent connectivity, clock skew between edge devices, and packet loss on MQTT brokers are the norm, not the exception. Design every data pipeline with an offline-first queue: buffer locally, transmit when connected, deduplicate on arrival. Never assume a reliable network.',
    tip:'This capstone maps directly onto your MES and Power Automate experience. The IoT telemetry handlers are the engineering equivalent of your Power Automate flows — event-triggered, stateful, and fault-tolerant. The difference is that here you are building the underlying infrastructure rather than configuring it. Use your MES domain knowledge to design realistic test scenarios — you already know which failure modes matter in production manufacturing.',
  },
];

for (const cap of capstones) {
  // Banner
  children.push(
    tbl([row([
      cell([
        p(cap.id,  {bold:true,color:'FFFFFF',size:22,align:AlignmentType.CENTER,after:10}),
        p(cap.badge,{bold:true,color:'FFFFFF',size:36,align:AlignmentType.CENTER,after:10}),
        p(cap.duration,{color:'D0DCF0',size:17,align:AlignmentType.CENTER,after:0}),
      ],{fill:cap.color,w:1440,bc:cap.color}),
      cell([
        p(cap.title, {bold:true,color:'FFFFFF',size:24,after:24}),
        p(cap.subtitle,{color:'A0C0E8',size:18,italic:true,after:0}),
      ],{fill:'1B2A4A',w:CW-1440,bc:'1B2A4A'}),
    ])],[1440,CW-1440]),
    sp(60),
    p(cap.stack,{size:20,color:'5A6A8A',italic:true,after:160}),
  );

  // What you're building
  children.push(h2('Platform Scale & Architecture', cap.color));
  for (const w of cap.what) children.push(bullet('',w));
  children.push(sp(120));

  // Component integrations table
  children.push(h2('Underlying Component Integrations', cap.color));
  children.push(tbl([
    row([
      cell([p('Component',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:cap.color,w:2200,bc:cap.color}),
      cell([p('Integration with Prior Projects',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:cap.color,w:CW-2200,bc:cap.color}),
    ]),
    ...cap.integrations.map((intg,i)=>row([
      cell([p(intg.label,{bold:true,size:19,color:cap.color,after:0})],{fill:i%2===0?cap.light:'FFFFFF',w:2200,bc:'CCCCCC'}),
      cell([p(intg.desc,{size:19,color:'333355',after:0})],{fill:i%2===0?cap.light:'FFFFFF',w:CW-2200,bc:'CCCCCC'}),
    ]))
  ],[2200,CW-2200]));
  children.push(sp(160));

  // Core concepts
  children.push(h2('Core Engineering Concepts', cap.color));
  for (const c of cap.concepts) children.push(num(c));
  children.push(sp(120));

  // Notes
  children.push(h2('Critical Implementation Notes', cap.color));
  children.push(noteBox('CAUTION', cap.hardPart,'7A5500','5A3A00','FFF8E0','E0C050'));
  children.push(sp(60));
  children.push(noteBox('TIP', cap.tip,'1A6A3A','0A4A2A','F0FFF8','50C080'));
  children.push(sp(120));

  // Deliverable note
  children.push(h2('Expected Outcome', cap.color));
  children.push(tbl([row([
    cell([p('Outcome: A fully integrated, deployed, and documented system demonstrating mastery of all prior project components working together as a production-grade platform.',{bold:true,size:20,color:'FFFFFF',after:0})],{fill:cap.color,w:CW,bc:cap.color}),
  ])],[CW]));
  children.push(sp(240), pb());
}

// Summary
children.push(
  h1('End State: Full Engineering Capability Matrix'),
  p('Every competency below is backed by a working, tested, deployed software artifact — not theoretical knowledge or tutorial completion.', {size:21,color:'333355',after:200}),
  tbl([
    row([
      cell([p('Domain',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:2000,bc:'1B2A4A'}),
      cell([p('Capability Acquired',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:4000,bc:'1B2A4A'}),
      cell([p('Proof Artifact',{bold:true,color:'FFFFFF',size:19,after:0})],{fill:'1B2A4A',w:3360,bc:'1B2A4A'}),
    ]),
    ...[
      ['Backend Infrastructure','Distributed caching, atomic concurrency, async queues, WebSocket streaming','P2 k6 load comparison chart and P3 chaos test 4-invariant proof'],
      ['AI and RAG Systems','Vector embedding, semantic retrieval, local LLM benchmarking pipeline','P1 benchmark dashboard and P4 live AI sentiment feature'],
      ['LLM Orchestration','Multi-agent pipelines, dependency chains, human-in-the-loop approval gates','P5 visual canvas Loom demo recording'],
      ['Transformer Architecture','Attention from raw tensors, causal masking, training loop, BPE tokenizer','P6 trained model with loss curve and attention heat maps'],
      ['Vector Indexing','IVF clustering, HNSW graph search, ANN recall-speed tradeoffs','P7 QPS vs recall benchmark chart and analytical report'],
      ['Model Quantization','INT4 compression, KV Cache, memory profiling, C++ dequant kernel','P8 4-panel telemetry dashboard'],
      ['Advanced Inference','MoE sparse routing, speculative decoding, throughput optimisation','P9 routing collapse ablation and P10 throughput benchmark dashboard'],
      ['ML Evaluation','Pass@K metrics, Docker sandboxed execution, LLM-as-judge, regression detection','P11 30-task evaluation suite with regression dashboard'],
      ['Fine-Tuning','LoRA PEFT, adapter merging, 6GB VRAM budget, synthetic dataset curation','P12 training curves and P11 improvement delta report'],
      ['Hardware Optimisation','Flash-Attention IO analysis, PCIe offloading, VRAM Floor Cap engineering','P13 IO comparison chart and P14 async pre-fetch dashboard'],
      ['Byte-Level Compression','Zstd dictionary training, error-bounded lossy quantization, streaming serialization','P15 throughput benchmark and compression ratio curve'],
      ['Distributed Scheduling','CAS lease acquisition, heartbeat zombie detection, dead-letter recovery, 1M-job test','P16 load test proof and orphan detection recording'],
      ['Tensor Parallelism','Column and row matrix sharding, ring-AllReduce, NCCL interconnect profiling','P17 4-panel interconnect dashboard and scaling efficiency chart'],
      ['Capstone Integration','Full-stack system synthesis combining all prior projects into deployed platforms','C1 Hydra Engine and C2 Industrial Copilot deployed with architecture docs'],
    ].map(([dom,cap,proof],i)=>row([
      cell([p(dom,{bold:true,size:19,color:'2E5090',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:2000,bc:'CCCCCC'}),
      cell([p(cap,{size:19,color:'1A1A2E',after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:4000,bc:'CCCCCC'}),
      cell([p(proof,{size:18,color:'5A6A8A',italic:true,after:0})],{fill:i%2===0?'EEF4FC':'FFFFFF',w:3360,bc:'CCCCCC'}),
    ]))
  ],[2000,4000,3360]),
  sp(200),
  p('This programme spans from backend application infrastructure through neural mathematics and hardware-level memory engineering, culminating in two capstone systems that integrate every prior project into production-grade deployments.', {size:21,color:'5A6A8A',italic:true,align:AlignmentType.CENTER,before:120})
);

// Assemble
const doc = new Document({
  numbering:{config:[
    {reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:480,hanging:240}},run:{font:'Arial'}}}]},
    {reference:'numbers',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:480,hanging:240}},run:{font:'Arial'}}}]},
  ]},
  styles:{
    default:{document:{run:{font:'Arial',size:22,color:'1A1A2E'}}},
    paragraphStyles:[
      {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:36,bold:true,font:'Arial',color:'1B2A4A'},paragraph:{spacing:{before:300,after:160},outlineLevel:0}},
      {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,font:'Arial',color:'2E5090'},paragraph:{spacing:{before:240,after:120},outlineLevel:1}},
    ],
  },
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:80},border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'2E5090',space:4}},children:[new TextRun({text:'Systems & AI Engineering Roadmap  ·  Arveend  ·  2025–2026',color:'8090B0',size:18,font:'Arial',italics:true})]})]})} ,
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:80,after:0},border:{top:{style:BorderStyle.SINGLE,size:4,color:'2E5090',space:4}},children:[new TextRun({text:'Page ',color:'8090B0',size:18,font:'Arial'}),            new TextRun({ text: " ", color: '8090B0', size: 18, font: 'Arial' }),new TextRun({text:'  ·  Self-Learn Engineering Programme',color:'8090B0',size:18,font:'Arial'})]})]})},
    children,
  }],
});

Packer.toBuffer(doc).then(b=>{fs.writeFileSync('/Users/ap/quiz/roadmap_plan.docx',b);console.log('Done');});