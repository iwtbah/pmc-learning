import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

// 运行：配套 Node.js 执行此文件。可用 ARTIFACT_NODE_MODULES 指定依赖目录。
// 所有坐标、字号均以 96 DPI 的像素计，输出画布为 1280 × 720（16:9）。
const outputDir = path.dirname(fileURLToPath(import.meta.url));
const workspaceDir = path.dirname(outputDir);
const runtimeRoot = '/Users/zhaowenzhuo/.cache/codex-runtimes/codex-primary-runtime/dependencies';
const moduleRoot = process.env.ARTIFACT_NODE_MODULES || path.join(runtimeRoot, 'node/node_modules');
process.env.RUNTIME_NODE_MODULES = moduleRoot;
const { Presentation, PresentationFile } = await import(pathToFileURL(path.join(moduleRoot, '@oai/artifact-tool/dist/artifact_tool.mjs')).href);
const skillDir = '/Users/zhaowenzhuo/.codex/plugins/cache/openai-primary-runtime/presentations/26.1015.12040/skills/presentations';
const buildDir = path.join(workspaceDir, '.codex-ppt-build');
await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(path.join(outputDir, '预览'), { recursive: true });

// 已核对本机字体。若后续环境安装了微软雅黑或苹方，可通过 PPT_FONT 替换。
const FONT = process.env.PPT_FONT || 'Hiragino Sans GB';
const C = { navy: '#101F35', blue: '#2487D0', light: '#69B9ED', bg: '#F3F6FA', white: '#FFFFFF', ink: '#243247', gray: '#65768B', line: '#D9E3ED', pale: '#EAF4FC', future: '#8C9DAF' };
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
deck.theme.defaultFont = FONT;
let serial = 0;
function shape(s, name, x, y, w, h, fill, border = 'none', radius = 0, geometry = 'rect') {
  return s.shapes.add({ geometry, name: `${name}-${++serial}`, position: { left: x, top: y, width: w, height: h }, fill,
    line: { fill: border, width: border === 'none' ? 0 : 1, style: 'solid' }, ...(radius ? { borderRadius: radius } : {}) });
}
function txt(s, value, x, y, w, h, size = 24, color = C.ink, bold = false, align = 'left') {
  const t = shape(s, 'text', x, y, w, h, 'none', 'none', 0, 'textbox');
  t.text = value;
  t.text.style = { fontSize: size, typeface: FONT, color, bold, alignment: align, verticalAlignment: 'middle', insets: 0, autoFit: 'none', wrap: 'none' };
  return t;
}
function line(s, x, y, w, h = 0, color = C.line, weight = 1) {
  const o = shape(s, 'line', x, y, w, h, 'none', color, 0, 'line');
  o.line = { fill: color, width: weight, style: 'solid' };
  return o;
}
function card(s, x, y, w, h, fill = C.white, border = C.line) { return shape(s, 'card', x, y, w, h, fill, border, 9); }
function footer(s, n, dark = false) {
  const col = dark ? '#9CB1CA' : C.gray;
  line(s, 64, 682, 1152, 0, dark ? '#30445F' : C.line);
  txt(s, 'PMC 阶段学习汇报', 64, 690, 400, 18, 16, col);
  txt(s, `${String(n).padStart(2, '0')} / 05`, 1116, 690, 100, 18, 16, col, false, 'right');
}
function page(n, title) {
  const s = deck.slides.add(); s.background.fill = C.bg;
  shape(s, 'title-accent', 64, 54, 6, 38, C.blue);
  txt(s, title, 88, 48, 1128, 52, 38, C.navy, true);
  footer(s, n); return s;
}
function notes(s, text) { s.speakerNotes.text = text; }

// 01 封面：原生文本及用户指定的抽象网格线，便于逐项编辑。
{
  const s = deck.slides.add(); s.background.fill = C.navy;
  shape(s, 'cover-accent', 72, 142, 66, 6, C.light);
  txt(s, 'PMC 业务学习阶段总结\n与 AI 技术应用探索', 72, 190, 900, 154, 52, C.white, true);
  txt(s, '阶段学习成果 · 业务理解 · AI 探索 · 后续规划', 76, 380, 850, 36, 26, '#B7CBE0');
  txt(s, '汇报人：[待填写]', 76, 516, 600, 32, 23, C.white);
  txt(s, '汇报时间：2026 年 10 月', 76, 558, 600, 32, 23, '#B7CBE0');
  for (let i = 0; i < 6; i++) line(s, 994 + i * 40, 160, 0, 390, '#203953');
  for (let i = 0; i < 10; i++) line(s, 970, 170 + i * 40, 260, 0, '#203953');
  line(s, 1010, 410, 80, 0, C.blue, 3); line(s, 1090, 330, 0, 80, C.blue, 3);
  line(s, 1090, 330, 120, 0, C.light, 3);
  for (const [x, y] of [[1006,406],[1086,326],[1206,326]]) shape(s, 'grid-node', x, y, 8, 8, C.light, 'none', 0, 'ellipse');
  footer(s, 1, true);
  notes(s, '汇报依据：工作区《学习目标与计划.md》《学习进度.md》，以及用户指定的阶段汇报大纲。业务背景为汽车线束教学模拟案例，不代表真实企业项目。汇报人为待填写项。');
}

// 02 学习进度：章节覆盖与掌握程度分开表达。
{
  const s = page(2, '阶段学习进度｜从基础认知向计划管理推进');
  const metrics = [['8 周', '总体学习周期'], ['15 章', '学习模块总数'], ['06  MRP', '当前学习位置']];
  metrics.forEach(([v, label], i) => {
    const x = 64 + 392 * i; card(s, x, 130, 368, 104);
    txt(s, v, x + 24, 144, 320, 44, 40, C.blue, true);
    txt(s, label, x + 24, 194, 320, 25, 21, C.gray);
  });
  txt(s, '已学习范围：第 01–06 章', 64, 256, 700, 34, 26, C.ink, true);
  txt(s, '6/15 章 · 40%', 886, 253, 330, 40, 30, C.blue, true, 'right');
  const gap = 6, sw = (1152 - 14 * gap) / 15;
  for (let i = 0; i < 15; i++) shape(s, `chapter-${i + 1}`, 64 + i * (sw + gap), 309, sw, 14, i < 5 ? C.blue : i === 5 ? C.light : '#D9E3ED', 'none', 3);
  txt(s, '40% 为章节覆盖率，不代表知识掌握度。第 4–6 章综合练习仍待完成。', 64, 337, 1152, 27, 18, C.gray);
  const stages = [
    ['阶段一', '基础认知与计划管理', '供应链基础、制造基础数据\n需求管理、S&OP、MPS、MRP', '已覆盖，持续复盘'],
    ['阶段二', '执行与资源管理', '产能、库存、采购、物料控制\n生产控制、排程', '后续学习'],
    ['阶段三', '综合应用', '精益生产、供应链系统、数据分析\n综合项目', '后续学习'],
  ];
  stages.forEach(([a, b, body, status], i) => {
    const x = 64 + i * 392; card(s, x, 390, 368, 177, i ? C.white : C.pale, i ? C.line : '#A9CEE9');
    txt(s, a, x + 22, 405, 110, 24, 18, i ? C.gray : C.blue, true);
    txt(s, status, x + 150, 405, 196, 24, 17, i ? C.future : C.blue, false, 'right');
    txt(s, b, x + 22, 440, 324, 31, 25, C.navy, true);
    txt(s, body, x + 22, 484, 328, 56, 20, C.gray);
  });
  txt(s, '学习方法', 64, 605, 136, 30, 22, C.navy, true);
  const methods = ['系统学习', '案例理解', '笔记整理', '阶段复盘'];
  methods.forEach((m, i) => {
    const x = 244 + i * 248; txt(s, m, x, 602, 172, 35, 24, C.ink, false, 'center');
    if (i < 3) shape(s, 'method-arrow', x + 194, 612, 24, 12, C.blue, 'none', 0, 'rightArrow');
  });
  txt(s, '贯穿案例：汽车线束制造（教学模拟）', 64, 650, 1152, 24, 17, C.gray);
  notes(s, '来源：《学习目标与计划.md》第3节课程安排；《学习进度.md》当前学习位置与各章总览。8周是总体规划，不代表已耗时8周。6÷15=40%，仅计算接触过的章节范围。当前为6.6，6.1～6.6框架已讲解，综合练习尚未开展。第4、5章整章验证亦未完成。阶段三的综合项目属于课程最终实践，不额外计入15章。S&OP：销售与运营计划（Sales & Operations Planning）；MPS：主生产计划（Master Production Schedule）；MRP：物料需求计划（Material Requirements Planning）。');
}

// 03 已学习内容：知识覆盖和阶段认知，避免宣称整章掌握。
{
  const s = page(3, '知识体系建立｜从业务认知到计划逻辑');
  txt(s, '已学习知识体系', 64, 131, 550, 36, 27, C.navy, true);
  txt(s, '阶段性认知收获', 670, 131, 546, 36, 27, C.navy, true);
  const left = [
    ['01–02', '基础认知', '供应链流程、PMC 职责、制造基础数据'],
    ['03–04', '需求与协同', '客户需求、预测、S&OP'],
    ['05–06', '计划管理', 'MPS、MRP、计划逻辑'],
  ];
  const right = [
    ['业务流程认知', '理解订单、计划、采购、生产、交付\n之间的关联'],
    ['基础数据认知', '认识 BOM、库存、提前期等基础数据\n对计划的影响'],
    ['系统性思维', '逐步建立从需求到计划再到执行的\n整体业务认识'],
  ];
  left.forEach(([index, title, body], i) => {
    const y = 188 + i * 130; card(s, 64, y, 550, 112);
    shape(s, 'level-band', 64, y, 94, 112, i === 2 ? C.blue : C.pale, 'none', 9);
    txt(s, index, 70, y + 38, 82, 36, 23, i === 2 ? C.white : C.blue, true, 'center');
    txt(s, title, 180, y + 17, 414, 35, 27, C.navy, true);
    txt(s, body, 180, y + 65, 418, 29, 21, C.gray);
    card(s, 670, y, 546, 112);
    txt(s, right[i][0], 694, y + 12, 498, 34, 26, C.navy, true);
    txt(s, right[i][1], 694, y + 51, 498, 52, 22, C.gray);
  });
  card(s, 64, 600, 1152, 63, C.navy, C.navy);
  txt(s, '从单点概念学习，逐步建立对 PMC 业务流程和计划协同逻辑的系统认识', 88, 611, 1104, 40, 25, C.white, true, 'center');
  notes(s, '来源：《学习目标与计划.md》第4节01～06模块；《学习进度.md》01～06章学习状态。此页表述学习范围和逐步形成的认知，不代表全部章节已完成独立验证或具备独立业务实操能力。贯穿案例为汽车线束制造教学模拟案例。PMC：生产与物料控制（Production & Material Control）；BOM：物料清单（Bill of Materials）；提前期（Lead Time）。');
}

// 04 AI：全部是探索方向，未作实际使用成效、掌握程度或落地收益声明。
{
  const s = page(4, 'AI 技术探索｜从工具使用到业务场景思考');
  txt(s, 'AI 工具学习方向', 64, 132, 545, 36, 27, C.navy, true);
  txt(s, '潜在业务应用方向', 680, 132, 536, 36, 27, C.navy, true);
  txt(s, '探索方向', 64, 173, 545, 24, 18, C.blue);
  txt(s, '待场景验证', 680, 173, 536, 24, 18, C.blue);
  const rows = [
    ['AI 知识辅助', '知识检索、归纳总结、学习辅助', '智能化辅助', '知识问答、流程辅助、信息提取'],
    ['AI 编程辅助', '代码生成、调试、数据处理', '数据处理', 'Excel 数据整理、自动化报表'],
    ['Agent 与工作流', 'Skill、MCP、工具调用、自动化探索', '业务分析', '库存、缺料、交付异常辅助分析'],
  ];
  rows.forEach(([title, body, target, desc], i) => {
    const y = 215 + i * 116;
    card(s, 64, y, 536, 100); card(s, 680, y, 536, 100, C.pale, '#C7E0F3');
    txt(s, title, 88, y + 12, 488, 34, 27, C.navy, true);
    txt(s, body, 88, y + 57, 488, 29, i === 2 ? 21 : 22, C.gray);
    shape(s, 'capability-application-arrow', 622, y + 44, 36, 14, C.blue, 'none', 0, 'rightArrow');
    txt(s, target, 704, y + 12, 488, 34, 27, C.navy, true);
    txt(s, desc, 704, y + 57, 488, 29, 22, C.gray);
  });
  txt(s, '工具示例：ChatGPT、Codex。以上为探索设想，未作为已落地成果或能力认证。', 64, 566, 1152, 26, 18, C.gray);
  card(s, 64, 612, 1152, 51, C.navy, C.navy);
  txt(s, '以实际业务问题为出发点，探索 AI、编程与数据分析工具的适用场景', 88, 620, 1104, 34, 25, C.white, true, 'center');
  notes(s, '内容依据：用户给定的AI学习与探索方向。工作区资料未提供实际AI学习程度、已完成AI业务项目或效果数据，因此全部标为探索与待验证方向。ChatGPT和Codex仅为用户提供的工具示例。左右箭头表示可能的关联，不表示唯一对应或已实现的系统集成。Agent指智能体；Skill指可复用技能指令；MCP指模型上下文协议（Model Context Protocol），用于连接工具与数据来源。');
}

// 05 双轨规划：阶段仅表示推进顺序，不虚构完成日期。
{
  const s = page(5, '下一阶段规划｜深化业务理解，探索技术赋能');
  txt(s, 'PMC 业务能力深化', 64, 132, 550, 36, 27, C.navy, true);
  txt(s, 'AI 技术应用探索', 680, 132, 536, 36, 27, C.navy, true);
  const labels = ['近期', '中期', '后期'];
  const business = ['产能、库存、采购、物料控制', '生产控制、排程、精益生产', '供应链系统、数据分析、综合项目'];
  const tech = ['深化 AI 工具与编程应用', '选择具体业务场景开展验证', '探索辅助分析工具或自动化方案'];
  for (const x of [95, 711]) line(s, x, 212, 0, 200, '#BED4E6', 3);
  labels.forEach((label, i) => {
    const y = 188 + i * 88;
    for (const [x, body] of [[64, business[i]], [680, tech[i]]]) {
      shape(s, 'timeline-node', x + 18, y + 22, 26, 26, i === 0 ? C.blue : C.light, C.bg, 0, 'ellipse');
      card(s, x + 62, y, 474, 72);
      txt(s, label, x + 80, y + 8, 410, 22, 17, C.blue, true);
      txt(s, body, x + 80, y + 34, 410, 28, 23, C.ink);
    }
  });
  txt(s, '推进安排：先完成第 4–6 章阶段复盘与综合练习，再按课程路线继续。', 64, 456, 1152, 26, 18, C.gray);
  txt(s, '预期交付成果', 64, 496, 1152, 28, 23, C.navy, true);
  const deliverables = ['系统化 PMC 学习笔记\n与知识框架', '汽车线束业务\n综合计划分析案例', 'AI 辅助数据处理\n或业务分析原型'];
  deliverables.forEach((d, i) => {
    const x = 64 + 392 * i; card(s, x, 535, 368, 72, C.pale, '#C7E0F3');
    txt(s, String(i + 1).padStart(2, '0'), x + 18, 551, 48, 38, 28, C.blue, true);
    txt(s, d, x + 82, 542, 267, 58, 22, C.ink);
  });
  txt(s, '以业务理解为基础、以数据分析为支撑、以 AI 技术为辅助，\n逐步探索专业知识向实际工作价值的转化', 64, 623, 1152, 51, 23, C.navy, true, 'center');
  notes(s, 'PMC课程路线来源：《学习目标与计划.md》第3节第4～8周课程与最终综合项目要求。阶段规划与AI探索方向来源于用户给定大纲。近期、中期、后期仅表示推进顺序，不承诺具体日期。综合计划分析案例、AI原型为预期交付，尚不作为已完成成果。第4～6章综合练习待完成的状态依据《学习进度.md》。业务案例将使用教学模拟数据，未引用真实企业生产或改善结果。');
}

const candidatePath = path.join(buildDir, 'PMC学习阶段汇报-候选.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
console.log('Editable draft exported');
for (const [i, s] of deck.slides.items.entries()) {
  const blob = await deck.export({ slide: s, format: 'png', scale: 1.5 });
  await fs.writeFile(path.join(outputDir, '预览', `第${i + 1}页.png`), new Uint8Array(await blob.arrayBuffer()));
  await fs.writeFile(path.join(buildDir, `slide-${i + 1}.layout.json`), await (await s.export({ format: 'layout' })).text());
  console.log(`Rendered slide ${i + 1}`);
}
// PDF 用逐页高分辨率预览合成，避免本地 PDF 渲染器的标准字体依赖。
// PPTX 中的文本和形状仍全部为原生可编辑对象。
const { PDFDocument } = await import(pathToFileURL(path.join(moduleRoot, 'pdf-lib/cjs/index.js')).href);
const pdf = await PDFDocument.create();
for (let i = 1; i <= 5; i++) {
  const image = await pdf.embedPng(await fs.readFile(path.join(outputDir, '预览', `第${i}页.png`)));
  const page = pdf.addPage([960, 540]);
  page.drawImage(image, { x: 0, y: 0, width: 960, height: 540 });
}
await fs.writeFile(path.join(outputDir, 'PMC学习阶段汇报-预览.pdf'), await pdf.save());
await fs.writeFile(path.join(buildDir, 'presentation.json'), JSON.stringify(deck.toProto()));

const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, 'container_tools/artifact_tool_utils.mjs')).href);
// 后续重新生成时保留原稿，自动另存为修订文件。
let finalPath = path.join(outputDir, 'PMC学习阶段汇报.pptx');
for (let revision = 2; await fs.access(finalPath).then(() => true).catch(() => false); revision++) {
  finalPath = path.join(outputDir, `PMC学习阶段汇报-修订${revision}.pptx`);
}
await finalizePresentation({
  explicitTotalSlideCount: 5,
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: path.join(runtimeRoot, 'python/bin/python3'),
  integrityValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit'],
  fontPolicy: { basis: 'design', families: [FONT] },
  requiredNativeTableOwnerSlides: [], requiredNativeChartOwnerSlides: [],
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, 'validation.json'),
});
console.log(`Final PPTX saved: ${finalPath}`);
