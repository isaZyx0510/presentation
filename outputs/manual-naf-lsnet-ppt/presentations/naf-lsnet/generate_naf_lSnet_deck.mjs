import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const artifactPath = "C:/Users/y50058176/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactPath).href);

const W = 1280;
const H = 720;
const OUT = "D:/presentation/NAFEncoderLSNet_architecture.pptx";
const PREVIEW_DIR = "D:/presentation/outputs/manual-naf-lsnet-ppt/presentations/naf-lsnet/preview";

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
  addText(slide, kicker, 54, 28, 620, 22, { size: 12, color: C.muted, bold: true });
  addText(slide, text, 54, 52, 900, 54, { size: 31, bold: true, face: "Microsoft YaHei UI" });
  addShape(slide, 54, 118, 112, 4, { fill: "#2563EB", line: line("#2563EB", 0) });
}

function footer(slide, n) {
  addText(slide, `NAFEncoderLSNet | ${String(n).padStart(2, "0")}`, 1042, 674, 178, 22, {
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
    addText(slide, it, x + 18, yy, w - 18, opts.h ?? 34, { size: opts.size ?? 14, color: opts.color ?? C.muted });
  });
}

const prs = Presentation.create({ slideSize: { width: W, height: H } });

function baseSlide(n) {
  const slide = prs.slides.add();
  slide.background.fill = C.bg;
  addShape(slide, 0, 0, W, H, { fill: C.bg, line: line(C.bg, 0) });
  footer(slide, n);
  return slide;
}

// 1
{
  const s = baseSlide(1);
  addText(s, "NAFEncoderLSNet 架构", 64, 88, 720, 70, { size: 40, bold: true, face: "Microsoft YaHei UI" });
  addText(s, "Encoder-only NAF backbone + LS recovery parameter heads", 68, 164, 780, 34, { size: 21, color: C.muted });
  const xs = [82, 315, 548, 781, 1014];
  const labels = [
    "Input adapter\nB x C_in x N_sc x T",
    "NAF Encoder\nstem + 2 downsamples",
    "Bottleneck\nB x 4C x N_sc/4 x T/4",
    "Heads\nmu / rho / delta / alpha",
    "Output contract\ntheta + alpha + path_prob",
  ];
  labels.forEach((l, i) => {
    box(s, l, xs[i], 318, 168, 96, i === 0 || i === 4 ? "input" : i === 1 ? "module" : i === 3 ? "head" : "process", { size: 15.5, bold: true });
    if (i < labels.length - 1) arrow(s, xs[i] + 174, 366, xs[i + 1] - 8, 366);
  });
  addText(s, "讲述主线：不预测 G_hat，而是从压缩后的 bottleneck 直接估计物理参数，并交给 recoveryMode='ls_from_packed_theta' 做 LS 恢复。", 88, 520, 1035, 58, { size: 17, color: C.muted });
}

// 2
{
  const s = baseSlide(2);
  title(s, "配置入口与物理范围", "Initialization");
  table(s, [
    ["变量", "来源", "用途"],
    ["n_rx", "cfg.N_rx", "接收天线数量，决定 delta_rx 维度"],
    ["l_eff", "cfg.L_eff", "有效路径数量，决定各 path head 维度"],
    ["n_sc", "cfg.N_sc", "子载波数量，决定频率轴 soft-argmax 分辨率"],
    ["in_ch", "input_adapter_dim(cfg)", "输入通道数，适配不同观测打包方式"],
    ["base_ch / hidden / dropout", "cfg or default", "控制 encoder 宽度、MLP 宽度和正则化"],
  ], 74, 154, 560, 54, [0.28, 0.31, 0.41], { size: 12.7 });
  box(s, "_slope_bounds(bounds, scale, sign)", 742, 172, 330, 72, "module", { size: 18, bold: true });
  arrow(s, 906, 248, 906, 292);
  box(s, "mu_min / mu_max\nrelDelayRangeSec + subcarrierSpacingHz\nsign = -1", 710, 300, 392, 72, "process", { size: 14.5, bold: true });
  box(s, "rho_min / rho_max\ndopplerRangeHz + ofdmSymbolDuration_s\nsign = +1", 710, 394, 392, 72, "process", { size: 14.5, bold: true });
  box(s, "delta_bound\nrxOffsetRangeSec 的最大绝对斜率", 710, 488, 392, 72, "process", { size: 14.5, bold: true });
  notes(s, [
    "所有 head 的输出先被约束到物理可解释范围，再做 gauge fix。",
    "mu 和 delta 使用负号映射延迟斜率，rho 使用正号映射 Doppler 斜率。",
  ], 76, 596, 780, { size: 14.5, gap: 30 });
}

// 3
{
  const s = baseSlide(3);
  title(s, "Encoder 主干", "Stem, downsampling, bottleneck");
  const y = 240;
  const nodes = [
    ["x\nB x C_in x N_sc x T", "input", 84],
    ["stem\nConv2d 1x1\nC_in -> C", "module", 300],
    ["stage1\nNAFBlock x1", "module", 510],
    ["down1\nConv3x3 s=2\nC -> 2C", "process", 720],
    ["stage2\nNAFBlock x2", "module", 930],
  ];
  nodes.forEach(([l, k, x], i) => {
    box(s, l, x, y, 156, 94, k, { size: 14.5, bold: true });
    if (i < nodes.length - 1) arrow(s, x + 162, y + 47, nodes[i + 1][2] - 8, y + 47);
  });
  box(s, "down2\nConv3x3 s=2\n2C -> 4C", 302, 430, 190, 90, "process", { size: 15, bold: true });
  box(s, "dropout2d\noptional regularization", 546, 430, 190, 90, "process", { size: 15, bold: true });
  box(s, "bottleneck\nNAFBlock x1\nB x 4C x N_sc/4 x T/4", 790, 430, 250, 90, "module", { size: 15, bold: true });
  arrow(s, 1008, 338, 397, 424, { color: "#64748B", width: 1.5 });
  arrow(s, 496, 475, 540, 475);
  arrow(s, 740, 475, 784, 475);
  notes(s, [
    "两次 stride=2 是各向同性下采样，频率轴和时间轴同时压缩。",
    "当 T=2 时，第一次下采样后时间轴通常已经变成 1，因此 rho 不走时间 soft-argmax。",
  ], 86, 594, 970, { size: 15, gap: 30 });
}

// 4
{
  const s = baseSlide(4);
  title(s, "mu Head: 频率轴 Soft-Argmax", "Path delay slope");
  box(s, "bottleneck feature f\nB x 4C x H' x T'", 90, 238, 220, 84, "process", { size: 17, bold: true });
  box(s, "mu_proj\nConv1x1: 4C -> L_eff", 390, 238, 230, 84, "head", { size: 17, bold: true });
  box(s, "mean over time\nB x L_eff x H'", 700, 238, 230, 84, "process", { size: 17, bold: true });
  box(s, "softmax + grid sum\nrange: [mu_min, mu_max]", 1000, 238, 230, 84, "gauge", { size: 16, bold: true });
  arrow(s, 316, 280, 384, 280);
  arrow(s, 626, 280, 694, 280);
  arrow(s, 936, 280, 994, 280);
  box(s, "mu_abs\nB x L_eff", 358, 454, 190, 70, "process", { size: 18, bold: true });
  box(s, "gauge_fix_mu(mu_abs)", 628, 454, 240, 70, "gauge", { size: 18, bold: true });
  box(s, "mu0, mu_rel\nreference + relative paths", 948, 454, 230, 70, "input", { size: 17, bold: true });
  arrow(s, 1114, 326, 454, 448, { color: "#64748B", width: 1.5 });
  arrow(s, 552, 489, 622, 489);
  arrow(s, 872, 489, 942, 489);
  addText(s, "核心思想：每条 path 在压缩频率轴 H' 上形成一个概率分布，用期望值把离散位置映射为连续的 mu_abs。", 90, 586, 980, 46, { size: 16, color: C.muted });
}

// 5
{
  const s = baseSlide(5);
  title(s, "rho / delta / alpha Heads", "Global pooled bottleneck");
  box(s, "f.mean(dim=(2,3))\npooled: B x 4C", 516, 148, 250, 74, "process", { size: 18, bold: true });
  const lanes = [
    ["rho_head\nLinear 4C -> hidden -> L_eff", "sigmoid scale\nrho_min + extent * sigmoid", "gauge_fix_rho\nrho0, rho_rel", 118, "head", "gauge"],
    ["delta_head\nLinear 4C -> hidden -> N_rx-1", "tanh scale\n[-delta_bound, delta_bound]", "prepend zero + gauge_fix_delta\ndelta_rx", 474, "head", "gauge"],
    ["alpha_head\nLinear 4C -> hidden -> L_eff", "softplus\npositive path strength", "alpha\nB x L_eff", 830, "head", "input"],
  ];
  lanes.forEach(([a, b, c, x, k1, k3]) => {
    box(s, a, x, 292, 260, 82, k1, { size: 14.5, bold: true });
    box(s, b, x, 414, 260, 82, "process", { size: 14.5, bold: true });
    box(s, c, x, 536, 260, 82, k3, { size: 14.5, bold: true });
    arrow(s, 641, 226, x + 130, 286, { color: "#64748B", width: 1.5 });
    arrow(s, x + 130, 378, x + 130, 408);
    arrow(s, x + 130, 500, x + 130, 530);
  });
  addText(s, "rho 选择全局池化 MLP，是因为当前 T=2 时 bottleneck 时间轴会坍缩为单点；alpha 用 softplus 保证非负。", 90, 640, 1000, 28, { size: 15, color: C.muted });
}

// 6
{
  const s = baseSlide(6);
  title(s, "path_prob Head 与兼容性", "Optional existence probability");
  box(s, "cfg.usePathProbHead", 112, 186, 260, 78, "module", { size: 20, bold: true });
  box(s, "True\npath_prob_head MLP\n4C -> hidden -> L_eff", 482, 154, 330, 96, "head", { size: 17, bold: true });
  box(s, "False\nones(B, L_eff)", 482, 304, 330, 78, "process", { size: 18, bold: true });
  box(s, "path_prob\nB x L_eff", 924, 232, 210, 78, "input", { size: 20, bold: true });
  arrow(s, 378, 225, 476, 202);
  arrow(s, 378, 225, 476, 342);
  arrow(s, 818, 202, 918, 262);
  arrow(s, 818, 342, 918, 286);
  notes(s, [
    "默认关闭时保持旧 checkpoint 可加载：模型没有 path_prob_head 权重，也不会出现 missing key。",
    "开启时每条 path 额外学习一个存在概率，经过 sigmoid 限制在 0 到 1。",
    "关闭时等价于 legacy 行为：所有 L_eff 条 path 都被认为存在。",
  ], 124, 472, 940, { size: 16, gap: 40 });
}

// 7
{
  const s = baseSlide(7);
  title(s, "Forward Pass 汇总", "Runtime data flow");
  const nodes = [
    ["shape check\nx.ndim == 4", 82, 190, "input"],
    ["encode\nstem -> stage1 -> down1 -> stage2 -> down2 -> bottleneck", 310, 190, "module"],
    ["pooled\nmean over H', T'", 638, 190, "process"],
    ["parameter heads\nmu, rho, delta, alpha, path_prob", 866, 190, "head"],
  ];
  nodes.forEach(([l, x, y, k], i) => {
    box(s, l, x, y, i === 1 ? 250 : 190, 86, k, { size: 15.5, bold: true });
    if (i < nodes.length - 1) arrow(s, x + (i === 1 ? 256 : 196), y + 43, nodes[i + 1][1] - 8, y + 43);
  });
  box(s, "theta dict\nmu0 / rho0 / delta_rx / mu_rel / rho_rel", 250, 396, 360, 100, "gauge", { size: 17, bold: true });
  box(s, "alpha\npath strengths", 704, 384, 190, 70, "input", { size: 18, bold: true });
  box(s, "path_prob\nexistence scores", 704, 488, 190, 70, "input", { size: 18, bold: true });
  arrow(s, 966, 280, 430, 390, { color: "#64748B", width: 1.5 });
  arrow(s, 966, 280, 798, 378, { color: "#64748B", width: 1.5 });
  arrow(s, 966, 280, 798, 482, { color: "#64748B", width: 1.5 });
  addText(s, "返回值严格匹配 detr_path_dict 的 modelOutputMode contract；forward 不输出 G_hat。", 82, 616, 980, 28, { size: 16, color: C.muted });
}

// 8
{
  const s = baseSlide(8);
  title(s, "输出接口与 LS Recovery", "Closing view");
  table(s, [
    ["Field", "Shape", "Meaning"],
    ["theta.mu0", "B x 1", "mu 的参考项"],
    ["theta.rho0", "B x 1", "rho 的参考项"],
    ["theta.delta_rx", "B x N_rx", "接收端相对 offset"],
    ["theta.mu_rel", "B x L_eff", "相对 delay slope"],
    ["theta.rho_rel", "B x L_eff", "相对 Doppler slope"],
    ["alpha", "B x L_eff", "非负路径强度"],
    ["path_prob", "B x L_eff", "路径存在概率或全 1 legacy fallback"],
  ], 86, 154, 620, 48, [0.28, 0.22, 0.5], { size: 13 });
  box(s, "recoveryMode\n'ls_from_packed_theta'", 802, 204, 300, 78, "module", { size: 18, bold: true });
  box(s, "No G_hat branch\nencoder estimates physical theta only", 802, 336, 300, 78, "note", { size: 17, bold: true, lineStyle: "dash" });
  box(s, "LS recovery\nuses packed theta downstream", 802, 468, 300, 78, "gauge", { size: 17, bold: true });
  arrow(s, 952, 286, 952, 330);
  arrow(s, 952, 418, 952, 462);
  notes(s, [
    "这页可作为代码讲解的收束：网络负责物理参数估计，恢复步骤交给 LS。",
    "重点强调 gauge-fixed theta 与 alpha/path_prob 是后续恢复的接口。",
  ], 90, 612, 820, { size: 14.5, gap: 30 });
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
