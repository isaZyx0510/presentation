import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const artifactPath = "C:/Users/y50058176/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactPath).href);

const W = 1280;
const H = 720;
const OUT = "D:/presentation/DETR_network_architecture_editable_slide2_updated.pptx";
const PREVIEW_DIR = "D:/presentation/outputs/manual-detr-ppt/presentations/detr-architecture/preview";

const C = {
  bg: "#F7F9FC",
  ink: "#111827",
  muted: "#5B6472",
  faint: "#E6EAF0",
  input: "#B7E4C7",
  inputStroke: "#1B4332",
  process: "#A9D6E5",
  processStroke: "#014F86",
  module: "#FFE08A",
  moduleStroke: "#B7791F",
  attn: "#CDB4DB",
  attnStroke: "#6D28D9",
  head: "#F8C4C4",
  headStroke: "#9B2226",
  gauge: "#F6B26B",
  gaugeStroke: "#B45309",
  note: "#EEF1F5",
  noteStroke: "#9AA3AF",
  white: "#FFFFFF",
};

function line(fill = C.faint, width = 2, style = "solid") {
  return { fill, width, style };
}

function addShape(slide, x, y, w, h, opts = {}) {
  return slide.shapes.add({
    geometry: opts.geometry || "rect",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: opts.fill ?? C.white,
    line: opts.line ?? line(C.faint, 1),
  });
}

function addText(slide, text, x, y, w, h, opts = {}) {
  const sh = addShape(slide, x, y, w, h, {
    fill: opts.fill ?? "#00000000",
    line: opts.line ?? line("#00000000", 0),
    geometry: opts.geometry,
    name: opts.name,
  });
  sh.text = text;
  sh.text.fontSize = opts.size ?? 24;
  sh.text.typeface = opts.face ?? "Microsoft YaHei";
  sh.text.color = opts.color ?? C.ink;
  sh.text.bold = Boolean(opts.bold);
  sh.text.alignment = opts.align ?? "left";
  sh.text.verticalAlignment = opts.valign ?? "top";
  sh.text.insets = opts.insets ?? { left: 2, right: 2, top: 2, bottom: 2 };
  return sh;
}

function title(slide, text, kicker = "") {
  addText(slide, kicker, 54, 28, 560, 22, { size: 12, color: C.muted, bold: true });
  addText(slide, text, 54, 52, 800, 54, { size: 31, bold: true, face: "Microsoft YaHei UI" });
  addShape(slide, 54, 118, 112, 4, { fill: "#2563EB", line: line("#2563EB", 0) });
}

function footer(slide, n) {
  addText(slide, `DETR architecture | ${String(n).padStart(2, "0")}`, 1050, 674, 170, 22, {
    size: 10,
    color: C.muted,
    align: "right",
  });
}

function box(slide, label, x, y, w, h, kind = "process", opts = {}) {
  const palette = {
    input: [C.input, C.inputStroke],
    process: [C.process, C.processStroke],
    module: [C.module, C.moduleStroke],
    attn: [C.attn, C.attnStroke],
    head: [C.head, C.headStroke],
    gauge: [C.gauge, C.gaugeStroke],
    note: [C.note, C.noteStroke],
    white: [C.white, C.faint],
  }[kind];
  const sh = addShape(slide, x, y, w, h, {
    geometry: opts.geometry || "roundRect",
    fill: opts.fill ?? palette[0],
    line: line(opts.stroke ?? palette[1], opts.lineWidth ?? 2, opts.lineStyle || "solid"),
    name: opts.name,
  });
  sh.text = label;
  sh.text.fontSize = opts.size ?? 15;
  sh.text.typeface = opts.face ?? "Microsoft YaHei";
  sh.text.color = opts.color ?? C.ink;
  sh.text.bold = Boolean(opts.bold);
  sh.text.alignment = opts.align ?? "center";
  sh.text.verticalAlignment = opts.valign ?? "middle";
  sh.text.insets = opts.insets ?? { left: 10, right: 10, top: 6, bottom: 6 };
  return sh;
}

function arrow(slide, x1, y1, x2, y2, opts = {}) {
  const color = opts.color ?? "#334155";
  const dx = x2 - x1;
  const dy = y2 - y1;
  if (Math.abs(dy) <= 8 && Math.abs(dx) >= 18) {
    const h = opts.thickness ?? 16;
    return addShape(slide, Math.min(x1, x2), y1 - h / 2, Math.abs(dx), h, {
      geometry: dx >= 0 ? "rightArrow" : "leftArrow",
      fill: color,
      line: line(color, 0),
    });
  }
  if (Math.abs(dx) <= 8 && Math.abs(dy) >= 18) {
    const w = opts.thickness ?? 16;
    return addShape(slide, x1 - w / 2, Math.min(y1, y2), w, Math.abs(dy), {
      geometry: dy >= 0 ? "downArrow" : "upArrow",
      fill: color,
      line: line(color, 0),
    });
  }
  return slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.max(Math.abs(dx), 1),
      height: Math.max(Math.abs(dy), 1),
    },
    line: line(color, opts.width ?? 2, opts.style || "solid"),
  });
}

function pill(slide, text, x, y, w, h, color = "#FFFFFF", stroke = C.faint, size = 13) {
  return box(slide, text, x, y, w, h, "white", {
    fill: color,
    stroke,
    size,
    geometry: "roundRect",
    lineWidth: 1.5,
  });
}

function table(slide, rows, x, y, w, rowH, cols, opts = {}) {
  const colXs = [x];
  for (let i = 0; i < cols.length - 1; i += 1) colXs.push(colXs[i] + cols[i] * w);
  for (let r = 0; r < rows.length; r += 1) {
    const fill = r === 0 ? opts.headerFill ?? "#E8EEF8" : r % 2 ? C.white : "#F4F7FB";
    for (let c = 0; c < cols.length; c += 1) {
      const cw = cols[c] * w;
      addShape(slide, colXs[c], y + r * rowH, cw, rowH, { fill, line: line(C.faint, 1) });
      addText(slide, rows[r][c], colXs[c] + 8, y + r * rowH + 5, cw - 16, rowH - 8, {
        size: opts.size ?? 13,
        bold: r === 0,
        color: r === 0 ? C.ink : C.muted,
        face: c === 0 && r > 0 ? "Aptos Mono" : "Microsoft YaHei",
      });
    }
  }
}

function notes(slide, items, x, y, w, opts = {}) {
  items.forEach((it, i) => {
    const yy = y + i * (opts.gap ?? 38);
    addShape(slide, x, yy + 5, 7, 7, { fill: opts.dot ?? "#2563EB", line: line(opts.dot ?? "#2563EB", 0), geometry: "ellipse" });
    addText(slide, it, x + 18, yy, w - 18, opts.h ?? 32, { size: opts.size ?? 14, color: opts.color ?? C.muted });
  });
}

const prs = Presentation.create({ slideSize: { width: W, height: H } });

function baseSlide(n, name) {
  const slide = prs.slides.add();
  slide.background.fill = C.bg;
  addShape(slide, 0, 0, W, H, { fill: C.bg, line: line(C.bg, 0) });
  footer(slide, n);
  return slide;
}

// 1
{
  const s = baseSlide(1);
  addText(s, "DETR 网络架构", 64, 88, 720, 70, { size: 44, bold: true, face: "Microsoft YaHei UI" });
  addText(s, "从观测张量到路径参数估计", 68, 164, 560, 34, { size: 22, color: C.muted });
  const xs = [92, 330, 568, 806, 1044];
  const labels = [
    "输入观测\nB x 2N_rx x N_sc x N_obs",
    "CNN stem\n局部特征 + 通道投影",
    "Transformer\nEncoder / Decoder attention",
    "Heads\nalpha / obj / mu / rho / delta",
    "输出参数\ntheta, alpha, path_prob, mix_lambda",
  ];
  labels.forEach((l, i) => {
    box(s, l, xs[i], 320, 160, 94, i === 0 || i === 4 ? "input" : i === 2 ? "attn" : i === 3 ? "head" : "module", { size: 16, bold: true });
    if (i < labels.length - 1) arrow(s, xs[i] + 166, 367, xs[i + 1] - 8, 367);
  });
  addText(s, "展示策略：先看浅层数据流，再展开 CNN stem、attention 和各参数 head；G head 当前不作为使用路径展开。", 88, 520, 980, 56, { size: 18, color: C.muted });
}

// 2
{
  const s = baseSlide(2);
  title(s, "整体浅层数据流", "主干视图");
  const y = 238;
  const nodes = [
    ["输入\nx + y", "input"],
    ["CNN stem\n局部特征 + 通道投影", "module"],
    ["Token + Pos\n二维 grid -> token 序列", "process"],
    ["Encoder\nself-attn over T", "attn"],
    ["Decoder\nq queries attend memory", "attn"],
    ["Heads\nalpha / obj / mu / rho / delta", "head"],
    ["Top-k + Gauge\n选择 L_eff 条路径", "gauge"],
    ["Outputs\ntheta + scores", "input"],
  ];
  const edgeLabels = [
    "x: B x 2N_rx x N_sc x N_obs\ny: training targets",
    "fmap: B x d_model x N_sc x N_obs",
    "fseq: B x T x d_model\nT = N_sc x N_obs",
    "memory: B x T x d_model",
    "decoded: B x q x d_model",
    "candidates: B x q",
    "selected: B x L_eff",
  ];
  nodes.forEach(([l, k], i) => {
    const x = 54 + i * 150;
    box(s, l, x, y, 128, 82, k, { size: 12.2, bold: true });
    if (i < nodes.length - 1) {
      arrow(s, x + 131, y + 41, x + 148, y + 41, { width: 2 });
      addText(s, edgeLabels[i], x + 101, y + 96, 102, 52, {
        size: 8.5,
        color: C.muted,
        align: "center",
      });
    }
  });
  box(s, "G head / G_hat / gains_all\n后续不使用，本稿不展开", 868, 512, 270, 74, "note", { size: 15, lineStyle: "dash" });
  notes(s, [
    "浅层流动只保留当前使用路径，避免把不使用的 G 分支误讲成核心输出。",
    "连接处标注该步骤的输入/输出维度：grid 表示进入 token 表示，再回到 L_eff 条路径参数。",
  ], 80, 510, 660, { size: 15 });
}

// 3
{
  const s = baseSlide(3);
  title(s, "输入与配置参数", "符号和物理口径");
  box(s, "Input x\nB x 2N_rx x N_sc x N_obs", 76, 202, 330, 118, "input", { size: 22, bold: true });
  notes(s, [
    "B: batch size",
    "2N_rx: 每个接收端复数观测拆成实部/虚部通道",
    "N_sc x N_obs: 子载波/观测符号二维网格",
    "shape check 不匹配时抛出 ValueError",
  ], 82, 360, 470, { size: 15, gap: 42 });
  table(s, [
    ["符号", "含义", "维度角色"],
    ["N_rx", "接收端数量", "输入通道和 delta_rx 输出"],
    ["N_tx", "发送端数量", "保留为配置符号，本稿不展开 G 分支"],
    ["N_sc", "子载波数量", "grid 高度 / pos_k 长度"],
    ["N_obs", "observedSymbols 长度", "grid 宽度 / pos_n 长度"],
    ["L_eff", "有效路径数", "top-k 输出路径数"],
    ["d_model", "Transformer 隐空间维度", "CNN stem 输出通道和 token 宽度"],
    ["q", "DETR query 数", "max(L_eff, detrNumQueries)"],
  ], 510, 168, 660, 46, [0.18, 0.42, 0.40], { size: 12.5 });
}

// 4
{
  const s = baseSlide(4);
  title(s, "Tensor Layout 与 Tokenization", "从二维观测网格到序列 token");
  box(s, "fmap\nB x d_model x N_sc x N_obs", 88, 198, 250, 90, "process", { size: 18, bold: true });
  box(s, "permute\nB x N_sc x N_obs x d_model", 420, 198, 260, 90, "process", { size: 17, bold: true });
  box(s, "reshape to tokens\nT = N_sc x N_obs\nfseq: B x T x d_model", 762, 190, 300, 106, "process", { size: 17, bold: true });
  arrow(s, 342, 243, 414, 243);
  arrow(s, 684, 243, 756, 243);
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c < 6; c += 1) {
      addShape(s, 134 + c * 30, 392 + r * 30, 24, 24, { fill: r % 2 ? "#D8ECF4" : "#B7DCEA", line: line("#7FB5CA", 1) });
    }
  }
  addText(s, "N_sc x N_obs grid", 122, 526, 210, 24, { size: 15, color: C.muted, align: "center" });
  arrow(s, 386, 454, 515, 454);
  for (let i = 0; i < 10; i += 1) {
    addShape(s, 548 + i * 34, 433, 28, 42, { fill: "#D8ECF4", line: line("#7FB5CA", 1) });
  }
  addText(s, "T tokens, each token width = d_model", 532, 526, 390, 24, { size: 15, color: C.muted, align: "center" });
  notes(s, ["Tokenization 不改变 batch B；只是把二维位置展开成 Transformer 可处理的序列。"], 870, 420, 300, { size: 15 });
}

// 5
{
  const s = baseSlide(5);
  title(s, "CNN Stem 设计", "局部特征提取与通道投影");
  box(s, "Input channels\n2N_rx", 95, 260, 170, 78, "input", { size: 20, bold: true });
  box(s, "Conv2d\n2N_rx -> d_model", 340, 246, 210, 106, "module", { size: 20, bold: true });
  box(s, "GELU", 595, 260, 116, 78, "module", { size: 21, bold: true });
  box(s, "Conv2d\nd_model -> d_model", 760, 246, 224, 106, "module", { size: 20, bold: true });
  box(s, "GELU\nfmap", 1035, 260, 126, 78, "module", { size: 21, bold: true });
  [265, 550, 711, 984].forEach((x, i) => arrow(s, x, 299, [336, 590, 756, 1030][i], 299));
  addText(s, "空间网格保持：N_sc x N_obs", 356, 390, 390, 28, { size: 18, color: C.muted });
  addText(s, "输出维度：B x d_model x N_sc x N_obs", 754, 390, 410, 28, { size: 18, color: C.muted });
  notes(s, [
    "第一层把复数接收通道映射到模型维度。",
    "第二层在 d_model 通道空间内继续提取局部模式。",
    "CNN stem 不展平 grid，为后续二维位置编码保留 N_sc 与 N_obs 结构。",
  ], 118, 492, 900, { size: 16, gap: 40 });
}

// 6
{
  const s = baseSlide(6);
  title(s, "二维位置编码与 Encoder", "把 grid 坐标注入 token 序列");
  box(s, "pos_k\nN_sc x d_model", 96, 202, 190, 70, "process", { size: 17, bold: true });
  box(s, "pos_n\nN_obs x d_model", 96, 310, 190, 70, "process", { size: 17, bold: true });
  box(s, "broadcast add\nN_sc x N_obs x d_model", 375, 246, 260, 92, "process", { size: 17, bold: true });
  box(s, "reshape + unsqueeze\npos: 1 x T x d_model", 710, 246, 240, 92, "process", { size: 17, bold: true });
  box(s, "encoder_input\nB x T x d_model", 474, 432, 220, 72, "input", { size: 17, bold: true });
  box(s, "Transformer Encoder\nself attention over T tokens", 790, 418, 280, 100, "attn", { size: 18, bold: true });
  box(s, "memory\nB x T x d_model", 850, 570, 184, 58, "process", { size: 16, bold: true });
  arrow(s, 290, 237, 369, 276);
  arrow(s, 290, 345, 369, 306);
  arrow(s, 638, 292, 704, 292);
  arrow(s, 830, 340, 678, 430);
  arrow(s, 695, 468, 784, 468);
  arrow(s, 930, 522, 930, 566);
  notes(s, ["位置编码是 learnable parameter；这里没有使用 _sincos_pos。"], 124, 514, 420, { size: 16 });
}

// 7
{
  const s = baseSlide(7);
  title(s, "Learnable Queries 与 Decoder Attention", "从 memory 中查询候选路径");
  box(s, "query parameters\nq x d_model", 92, 190, 220, 78, "process", { size: 18, bold: true });
  box(s, "unsqueeze + expand\nqueries: B x q x d_model", 92, 330, 238, 90, "process", { size: 17, bold: true });
  arrow(s, 202, 270, 202, 326);
  box(s, "memory\nB x T x d_model", 482, 190, 220, 78, "process", { size: 18, bold: true });
  box(s, "Decoder self-attention\nqueries exchange information", 452, 330, 286, 86, "attn", { size: 17, bold: true });
  box(s, "Cross-attention\nqueries attend encoder memory", 792, 330, 300, 86, "attn", { size: 17, bold: true });
  box(s, "decoded\nB x q x d_model", 850, 506, 210, 76, "process", { size: 18, bold: true });
  arrow(s, 332, 375, 446, 375);
  arrow(s, 704, 230, 908, 326);
  arrow(s, 740, 375, 786, 375);
  arrow(s, 942, 420, 942, 502);
  addText(s, "q 是候选路径槽位数；每个槽位最终会通过 heads 预测路径相关参数。", 92, 544, 620, 44, { size: 17, color: C.muted });
}

// 8
{
  const s = baseSlide(8);
  title(s, "路径存在性与基础参数 Heads", "decoded -> alpha / objectness / rho");
  box(s, "decoded\nB x q x d_model", 72, 286, 210, 94, "process", { size: 19, bold: true });
  const heads = [
    ["alpha_head\nLinear d_model -> 1\nsqueeze + softplus", "alpha_all\nB x q", 190],
    ["obj_head\nLinear d_model -> 1\nsqueeze + sigmoid", "path_prob_all\nB x q", 320],
    ["rho_head\nLinear d_model -> 1\nsqueeze + sigmoid + scale", "rho_raw_all\nB x q", 450],
  ];
  heads.forEach(([h, out, y]) => {
    box(s, h, 410, y, 260, 84, "head", { size: 15.5, bold: true });
    box(s, out, 820, y + 6, 210, 72, "process", { size: 17, bold: true });
    arrow(s, 286, 333, 404, y + 42, { color: "#64748B", width: 1.5 });
    arrow(s, 674, y + 42, 814, y + 42);
  });
  notes(s, [
    "alpha_all 表示路径幅度/权重的非负预测。",
    "path_prob_all 表示候选 query 成为有效路径的概率。",
    "rho_raw_all 是路径位置参数 rho 的 query 级预测。",
  ], 118, 560, 920, { size: 16, gap: 34 });
}

// 9
{
  const s = baseSlide(9);
  title(s, "mu Head 三种模式", "free / scaled / grid");
  box(s, "decoded\nB x q x d_model", 520, 144, 220, 70, "process", { size: 18, bold: true });
  const lanes = [
    ["free", "mu_head\nB x q x 1\nsqueeze -> B x q\nsigmoid + scale", "mu_raw_all\nB x q", 72],
    ["scaled", "mu_head\nB x q x 1\nsigmoid -> frac B x q", "frac + extent\nmu_min + frac * extent range", 462],
    ["grid", "mu_head\nB x q x num_mu_bins\nsoftmax over bins", "weighted sum with mu_dict\nfrac B x q", 852],
  ];
  lanes.forEach(([name, a, b, x]) => {
    addText(s, name, x, 244, 260, 28, { size: 22, bold: true, align: "center" });
    box(s, a, x, 286, 280, 112, "head", { size: 15.5, bold: true });
    box(s, b, x, 436, 280, 82, "process", { size: 15.5, bold: true });
    arrow(s, 630, 218, x + 140, 280);
    arrow(s, x + 140, 402, x + 140, 432);
  });
  box(s, "cross_path_stats\nmean over query dim\nB x d_model", 456, 564, 230, 74, "process", { size: 14.5, bold: true });
  box(s, "s_head MLP\nB x d_model -> B x 1\nextent", 720, 564, 230, 74, "head", { size: 14.5, bold: true });
  arrow(s, 686, 601, 714, 601);
  addText(s, "scaled/grid 共用 extent；free 分支直接给出 mu_raw_all。", 82, 634, 480, 24, { size: 14, color: C.muted });
}

// 10
{
  const s = baseSlide(10);
  title(s, "Top-k 选择与 Gauge Fix", "从 q 个候选槽位到 L_eff 条有效路径");
  box(s, "alpha_all\nB x q", 86, 172, 170, 60, "process", { size: 17, bold: true });
  box(s, "path_prob_all\nB x q", 86, 262, 170, 60, "process", { size: 17, bold: true });
  box(s, "score = alpha_all * path_prob_all\nB x q", 346, 212, 270, 82, "process", { size: 17, bold: true });
  box(s, "top-k over query dim\ntop_idx: B x L_eff", 704, 212, 250, 82, "gauge", { size: 17, bold: true });
  arrow(s, 260, 202, 340, 246);
  arrow(s, 260, 292, 340, 258);
  arrow(s, 620, 253, 698, 253);
  const gathers = [
    ["mu_raw_all -> mu_raw\nB x L_eff", 112],
    ["rho_raw_all -> rho_raw\nB x L_eff", 354],
    ["alpha_all -> alpha\nB x L_eff", 596],
    ["path_prob_all -> path_prob\nB x L_eff", 838],
  ];
  gathers.forEach(([l, x]) => {
    box(s, l, x, 410, 210, 64, "process", { size: 14.5, bold: true });
    arrow(s, 828, 298, x + 105, 406, { color: "#64748B", width: 1.5 });
  });
  box(s, "gauge_fix_mu\nmu0 = max(mu_raw)\nmu_rel = mu_raw - mu0", 174, 552, 270, 86, "gauge", { size: 15.5, bold: true });
  box(s, "gauge_fix_rho\nrho0 = mean(rho_raw)\nrho_rel = rho_raw - rho0", 506, 552, 270, 86, "gauge", { size: 15.5, bold: true });
  box(s, "delta path\nfmap -> pool -> delta_mlp\nB x N_rx -> relative offsets", 836, 552, 280, 86, "gauge", { size: 15.5, bold: true });
}

// 11
{
  const s = baseSlide(11);
  title(s, "输出接口与讲解收束", "当前使用的 forward path");
  box(s, "theta dict", 90, 160, 260, 64, "input", { size: 22, bold: true });
  table(s, [
    ["字段", "形状", "说明"],
    ["mu0", "B x 1", "路径位置参数 mu 的参考量"],
    ["rho0", "B x 1", "路径位置参数 rho 的参考量"],
    ["delta_rx", "B x N_rx", "接收端相对偏移"],
    ["mu_rel", "B x L_eff", "有效路径的相对 mu"],
    ["rho_rel", "B x L_eff", "有效路径的相对 rho"],
  ], 90, 240, 520, 48, [0.24, 0.24, 0.52], { size: 13.5 });
  box(s, "alpha\nB x L_eff", 690, 202, 190, 74, "process", { size: 19, bold: true });
  box(s, "path_prob\nB x L_eff", 930, 202, 190, 74, "process", { size: 19, bold: true });
  box(s, "mix_lambda\n1", 810, 330, 190, 74, "process", { size: 19, bold: true });
  box(s, "G head / G_hat\n当前展示版本不作为使用路径展开", 720, 500, 360, 82, "note", { size: 17, lineStyle: "dash" });
  notes(s, [
    "pack_slope_theta 当前 forward path 未使用，模型返回 dict。",
    "本稿主线到 theta、alpha、path_prob、mix_lambda 为止。",
  ], 92, 600, 700, { size: 15, gap: 30 });
}

await fs.mkdir(PREVIEW_DIR, { recursive: true });
for (let i = 0; i < prs.slides.count; i += 1) {
  const slide = prs.slides.getItem(i);
  const png = await prs.export({ slide, format: "png", scale: 1 });
  const buf = Buffer.from(await png.arrayBuffer());
  await fs.writeFile(path.join(PREVIEW_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`), buf);
}

const pptx = await PresentationFile.exportPptx(prs);
await pptx.save(OUT);
console.log(JSON.stringify({ output: OUT, slides: prs.slides.count, previewDir: PREVIEW_DIR }, null, 2));
