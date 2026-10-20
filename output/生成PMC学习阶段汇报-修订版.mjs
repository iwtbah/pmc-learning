import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

// 以配套 Node.js 运行。PPTX 的文字、流程图、进度条和表格均为原生对象。
const outputDir = path.dirname(fileURLToPath(import.meta.url));
const workspaceDir = path.dirname(outputDir);
const runtime = '/Users/zhaowenzhuo/.cache/codex-runtimes/codex-primary-runtime/dependencies';
const modules = process.env.ARTIFACT_NODE_MODULES || path.join(runtime, 'node/node_modules');
process.env.RUNTIME_NODE_MODULES = modules;
const { GlobalFonts } = await import(pathToFileURL(path.join(modules, '@napi-rs/canvas/index.js')).href);
// 预览也使用真实的微软雅黑，避免仅修改 PPTX 字体名称而使用替代字体渲染。
const fontDir = process.env.PPT_FONT_DIR || '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts';
for (const name of ['msyh.ttc', 'msyhbd.ttc']) GlobalFonts.registerFromPath(path.join(fontDir, name));
const { Presentation, PresentationFile } = await import(pathToFileURL(path.join(modules, '@oai/artifact-tool/dist/artifact_tool.mjs')).href);
const skill = '/Users/zhaowenzhuo/.codex/plugins/cache/openai-primary-runtime/presentations/26.1015.12040/skills/presentations';
const build = path.join(workspaceDir, '.codex-ppt-build', 'revision-2');
const preview = path.join(outputDir, '修订版预览');
await fs.mkdir(build, { recursive: true });
await fs.mkdir(preview, { recursive: true });

/* 设计：
 * 海军蓝 #101F35；科技蓝 #2487D0；浅蓝 #EAF3FA；正文 #243247；白底。
 * 字体：微软雅黑；标题 27pt，正文 14～15pt，辅助文字 11～12pt。
 * 页2：左侧周期和当前位置 / 右侧章节路线。
 * 页3：左侧完整计划业务链 / 右侧业务、数据、管理逻辑认知。
 * 页4：窄栏工具方向 / 宽栏候选应用场景原生表格。
 * 页5：两条学习主线按阶段对照 / 下方预期交付清单。
 * 不使用重复卡片、底部口号和装饰性标签。各页统一标题基线和页码。
 */
const C = { navy: '#101F35', blue: '#2487D0', pale: '#EAF3FA', white: '#FFFFFF', ink: '#243247', gray: '#687584', line: '#DCE4EB', muted: '#F4F6F8', pending: '#8D9BA9' };
const FONT = 'Microsoft YaHei';
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
deck.theme.defaultFont = FONT;
let serial = 0;
function obj(s, x, y, w, h, fill = 'none', geometry = 'rect', border = 'none', radius = 0) {
  return s.shapes.add({ name: `element-${++serial}`, geometry, position: { left: x, top: y, width: w, height: h }, fill,
    line: { fill: border, width: border === 'none' ? 0 : 1, style: 'solid' }, ...(radius ? { borderRadius: radius } : {}) });
}
function text(s, value, x, y, w, h, size = 20, color = C.ink, bold = false, align = 'left') {
  const t = obj(s, x, y, w, h, 'none', 'textbox'); t.text = value;
  t.text.style = { typeface: FONT, fontSize: size, color, bold, alignment: align, verticalAlignment: 'middle', insets: 0, wrap: 'none', autoFit: 'none' };
  return t;
}
function rule(s, x, y, w, h = 0, color = C.line, width = 1) {
  const r = obj(s, x, y, w, h, 'none', 'line'); r.line = { fill: color, width, style: 'solid' }; return r;
}
function page(n, title) {
  const s = deck.slides.add(); s.background.fill = C.white;
  text(s, title, 64, 50, 1152, 52, 36, C.navy, true);
  rule(s, 64, 116, 1152);
  text(s, String(n).padStart(2, '0'), 1166, 676, 50, 22, 16, C.pending, false, 'right');
  return s;
}
function note(s, v) { s.speakerNotes.text = v; }
function table(s, x, y, widths, heights, values, size = 20) {
  const t = s.tables.add({ rows: values.length, columns: widths.length, left: x, top: y,
    width: widths.reduce((a,b)=>a+b,0), height: heights.reduce((a,b)=>a+b,0), columnWidths: widths, values });
  t.styleOptions = { headerRow: false, bandedRows: false, firstColumn: false, lastColumn: false };
  t.borders.assign({ fill: C.line, width: 1, style: 'solid' });
  t.cells.block({ row: 0, column: 0, rowCount: values.length, columnCount: widths.length }).assign({ margins: { left: 20, right: 18, top: 12, bottom: 12 } });
  heights.forEach((h, r) => { t.rows[r].height = h; });
  values.forEach((row,r)=>row.forEach((_,c)=>{
    const cell=t.getCell(r,c); cell.fill = r===0 ? C.pale : C.white;
    cell.text.style={ typeface: FONT, fontSize: size, color: r===0 ? C.navy : C.ink, bold: r===0 || c===0,
      alignment: 'left', verticalAlignment: 'middle', autoFit: 'none', wrap: 'none' };
  }));
  return t;
}

// 1. 封面
{
  const s = deck.slides.add(); s.background.fill = C.navy;
  text(s, 'PMC 业务学习阶段总结\n与 AI 技术应用探索', 76, 208, 1000, 142, 46, C.white, true);
  rule(s, 78, 394, 82, 0, '#69B9ED', 3);
  text(s, '汇报人：[待填写]', 78, 510, 700, 32, 21, C.white);
  text(s, '2026 年 10 月', 78, 552, 700, 30, 20, '#B8C9DA');
  // 保留原要求中的简洁网格，压低对比度与面积。
  for(let i=0;i<5;i++) rule(s, 1040+i*34, 176, 0, 330, '#20364E');
  for(let i=0;i<9;i++) rule(s, 1014, 186+i*36, 202, 0, '#20364E');
  rule(s, 1040, 366, 102, 0, '#3B7AA8', 2);
  text(s, '01', 1166, 676, 50, 22, 16, '#8FA6BF', false, 'right');
  note(s, '依据：工作区《学习目标与计划.md》《学习进度.md》及用户阶段汇报大纲。汇报人为待填写项。业务案例属于教学模拟，不代表真实企业项目。');
}

// 2. 学习进度：去掉案例说明、学习方法装饰流程和阶段汇报页脚。
{
  const s=page(2, '学习计划与阶段进度');
  rule(s, 344, 160, 0, 464);
  text(s, '8 周', 64, 160, 248, 64, 46, C.navy, true);
  text(s, '总体学习周期', 64, 236, 248, 30, 19, C.gray);
  text(s, '15 章', 64, 310, 248, 64, 46, C.navy, true);
  text(s, '课程模块', 64, 386, 248, 30, 19, C.gray);
  text(s, '第 6 章 MRP', 64, 477, 248, 40, 27, C.blue, true);
  text(s, '当前学习位置\n框架已讲解，综合练习待完成', 64, 532, 254, 60, 18, C.gray);

  text(s, '章节覆盖', 404, 153, 600, 30, 22, C.navy, true);
  text(s, '6 / 15 章', 1044, 147, 172, 40, 27, C.blue, true, 'right');
  obj(s, 404, 211, 812, 9, C.line); obj(s, 404, 211, 812*0.4, 9, C.blue);
  text(s, '40% 为章节覆盖率，不代表知识掌握度', 404, 236, 812, 26, 16, C.gray);

  rule(s, 412, 323, 0, 274, '#D4DEE7', 2);
  const stages=[
    ['01–06', '基础认知与计划管理', '供应链基础、制造基础数据、需求管理\nS&OP、MPS、MRP', C.blue],
    ['07–12', '执行与资源管理', '产能、库存、采购、物料控制\n生产控制、排程', C.pending],
    ['13–15', '综合应用', '精益生产、供应链系统、数据分析及综合项目', C.pending],
  ];
  stages.forEach(([range,title,body,color],i)=>{
    const y=302+i*120;
    obj(s, 406, y+14, 12, 12, color, 'ellipse');
    text(s, range, 444, y, 104, 34, 21, color, true);
    text(s, title, 570, y, 646, 34, 23, i===0?C.navy:C.gray, true);
    text(s, i===0?'已覆盖':'后续学习', 1080, y+3, 136, 28, 16, color, false, 'right');
    text(s, body, 570, y+46, 646, 56, 19, C.gray);
  });
  note(s, '来源：《学习目标与计划.md》第3节课程安排；《学习进度.md》当前学习位置及各章进度总览。8周为计划周期，未作为实际耗时。6÷15=40%仅表示已经覆盖的章节范围。当前6.1～6.6框架已讲解，综合练习尚未开展，第4、5章综合验证也未完成。01～06为已学习范围，07～15为后续学习。综合项目不另计章节。S&OP为销售与运营计划，MPS为主生产计划，MRP为物料需求计划。');
}

// 3. 按用户图片的业务逻辑重排，流程图及上下游关系可编辑。
{
  const s=page(3, '已学知识体系与核心收获');
  rule(s, 824, 152, 0, 492);
  text(s, '基础认知（01–02）', 64, 144, 236, 30, 21, C.navy, true);
  text(s, '供应链流程、PMC 职责、制造基础数据', 310, 144, 484, 30, 18, C.gray);
  const nodes=[
    ['客户预测 / 正式订单', '需求管理：区分预测需求和确定性需求'],
    ['S&OP  销售与运营计划', '协调需求、供应、库存与产能'],
    ['MPS  主生产计划', '确定成品生产数量和时间'],
    ['MRP  物料需求计划', '结合 BOM、库存、在途与提前期形成物料需求'],
  ];
  nodes.forEach(([title,desc],i)=>{
    const y=208+i*90;
    obj(s, 64, y, 730, 68, i===3?C.pale:C.muted, 'rect', 'none', 5);
    text(s, title, 88, y+7, 682, 29, 22, C.navy, true, 'center');
    text(s, desc, 88, y+38, 682, 23, 18, C.gray, false, 'center');
    if(i<3) obj(s, 423, y+76, 12, 9, '#8B99A6', 'downArrow');
  });
  text(s, '后续业务', 64, 568, 730, 24, 16, C.gray, false, 'center');
  rule(s, 429, 552, 0, 10, '#AAB6C2');
  rule(s, 185, 600, 488, 0, '#AAB6C2');
  rule(s, 429, 592, 0, 8, '#AAB6C2');
  ['采购计划','生产执行','物料控制'].forEach((label,i)=>{
    const x=64+i*246; rule(s, x+119, 600, 0, 11, '#AAB6C2');
    obj(s, x, 611, 238, 39, C.muted, 'rect', 'none', 4);
    text(s, label, x+12, 615, 214, 30, 19, C.ink, false, 'center');
  });
  text(s, '核心收获', 866, 144, 350, 34, 24, C.navy, true);
  const gains=[
    ['业务认知', 'PMC 需要协同管理需求与计划，\n并统筹物料和资源。'],
    ['数据认知', 'BOM、库存、提前期等数据，\n直接影响计划运算结果。'],
    ['管理逻辑', '计划需要兼顾需求满足、\n交付风险与资源约束。'],
  ];
  gains.forEach(([title,body],i)=>{
    const y=220+i*134;
    text(s, title, 866, y, 350, 31, 22, C.blue, true);
    text(s, body, 866, y+45, 350, 61, 19, C.ink);
  });
  note(s, '来源：《学习目标与计划.md》01～06章与《学习进度.md》已讲范围。结构参考用户提供图片中的需求管理、S&OP、MPS、MRP与执行业务链。01～02提供业务和数据基础，03～06沿计划链展开。采购计划、生产执行和物料控制仅表示后续业务关系，不代表相应章节已经学习。认知收获为阶段理解，不代表完整独立操作能力。PMC：生产与物料控制（Production & Material Control）；BOM：物料清单（Bill of Materials）；提前期（Lead Time）。');
}

// 4. 工具方向与候选场景不逐项强行配对，业务提效针对重复工作流程。
{
  const s=page(4, 'AI 工具学习与业务应用探索');
  text(s, '工具学习方向', 64, 152, 308, 34, 24, C.navy, true);
  rule(s, 394, 152, 0, 460);
  text(s, 'AI 知识辅助', 64, 229, 308, 32, 22, C.blue, true);
  text(s, '知识检索、归纳总结\n学习辅助', 64, 280, 308, 60, 19, C.ink);
  text(s, '工具示例：ChatGPT、Codex', 64, 355, 308, 25, 16, C.gray);
  text(s, 'Agent 与工作流', 64, 444, 308, 32, 22, C.blue, true);
  text(s, 'Skill、MCP、工具调用\n自动化流程探索', 64, 495, 308, 60, 19, C.ink);

  text(s, '潜在业务应用', 438, 152, 778, 34, 24, C.navy, true);
  text(s, '候选场景，待业务验证', 438, 196, 778, 25, 16, C.gray);
  table(s,438,244,[160,618],[48,78,78,78,90],[
    ['应用方向','候选场景'],
    ['数据处理','Excel 数据整理、数据清洗与报表生成'],
    ['业务分析','库存、缺料与交付异常的辅助分析'],
    ['智能化辅助','知识问答、流程指引与信息提取'],
    ['业务提效','针对重复工作流程，探索资料自动汇总、\n核对与流转'],
  ],19);
  note(s, '依据：用户提供的AI探索方向。AI工具学习及业务应用均为探索方向，未宣称已经掌握、落地或取得效率收益。业务提效针对重复工作流程，资料汇总、核对、流转均属于候选场景，不作为已完成案例。Skill指可复用技能，MCP指模型上下文协议（Model Context Protocol），Agent指智能体。');
}

// 5. 按阶段对照两条主线，不使用重复时间线与底部口号。
{
  const s=page(5,'下一阶段学习与工作规划');
  table(s,64,166,[134,509,509],[50,104,88,88],[
    ['阶段','PMC 业务学习','AI 应用探索'],
    ['近期','完成阶段练习与复盘\n学习产能、库存、采购与物料控制','深化 AI 工具与工作流应用'],
    ['中期','生产控制、排程、精益生产','选择具体业务场景开展验证'],
    ['后期','供应链系统、数据分析、综合项目','探索辅助分析工具或自动化方案'],
  ],20);
  text(s, '预期交付', 64, 546, 1152, 34, 24, C.navy, true);
  const outputs=[['学习资料','PMC 学习笔记与知识框架'],['综合案例','汽车线束业务综合计划分析'],['应用原型','AI 辅助分析或流程自动化原型']];
  outputs.forEach(([label,body],i)=>{
    const x=64+i*392;
    if(i) rule(s,x-22,599,0,53);
    text(s,label,x,598,368,26,17,C.blue,true);
    text(s,body,x,637,368,30,19,C.ink);
  });
  note(s, 'PMC路线来源：《学习目标与计划.md》第3节课程安排与最终综合项目要求。第4～6章仍有待完成练习，近期先完成阶段复盘。AI应用探索及原型为未来计划，不表示已完成成果，不承诺具体日期或业务收益。预期交付的综合案例使用教学模拟数据。');
}

const candidatePath=path.join(build,'候选.pptx');
await (await PresentationFile.exportPptx(deck)).save(candidatePath);
for(const [i,s] of deck.slides.items.entries()){
  const b=await deck.export({slide:s,format:'png',scale:1.5});
  await fs.writeFile(path.join(preview,`第${i+1}页.png`),new Uint8Array(await b.arrayBuffer()));
  await fs.writeFile(path.join(build,`slide-${i+1}.layout.json`),await (await s.export({format:'layout'})).text());
  console.log(`Rendered ${i+1}`);
}
const { PDFDocument }=await import(pathToFileURL(path.join(modules,'pdf-lib/cjs/index.js')).href);
const pdf=await PDFDocument.create();
for(let i=1;i<=5;i++){
  const img=await pdf.embedPng(await fs.readFile(path.join(preview,`第${i}页.png`)));
  pdf.addPage([960,540]).drawImage(img,{x:0,y:0,width:960,height:540});
}
await fs.writeFile(path.join(outputDir,'PMC学习阶段汇报-修订版预览.pdf'),await pdf.save());
await fs.writeFile(path.join(build,'presentation.json'),JSON.stringify(deck.toProto()));
let finalPath=path.join(outputDir,'PMC学习阶段汇报-修订版.pptx');
for(let v=3;await fs.access(finalPath).then(()=>true).catch(()=>false);v++)finalPath=path.join(outputDir,`PMC学习阶段汇报-修订${v}.pptx`);
const { finalizePresentation }=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
await finalizePresentation({workspaceDir,candidatePath,finalPath,explicitTotalSlideCount:5,
  pythonExecutable:path.join(runtime,'python/bin/python3'),
  integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit','--validate-bullet-geometry','--require-native-table-slide','4','--require-native-table-slide','5'],
  requiredNativeTableOwnerSlides:[4,5],requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:'user_request',families:[FONT]},verifyArtifactToolImport:true,
  receiptPath:path.join(build,`validation-${Date.now()}.json`)});
console.log(`Saved: ${finalPath}`);
