import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const args = process.argv.slice(2);
const valueAfter = (flag) => {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : undefined;
};

const inputPath = valueAfter("--input");
const outputPath = valueAfter("--output") || "book-topic-evaluation.xlsx";
const verifyDir = valueAfter("--verify-dir");
const input = inputPath ? JSON.parse(await fs.readFile(inputPath, "utf8")) : {};

const dimensions = [
  {
    type: "必须项", name: "文化母体", weight: 16,
    question: "背后是否有足够强的共同认知、时代情绪、经典 IP 或社会议题？",
    risk: "需要反复解释为什么要关心它。",
    description: "判断是否有读者无需教育即可理解、认同或争论的共同语境。",
    high: "读者不需要大量解释就知道它和自己有关。",
    low: "需要反复解释“为什么要关心它”。",
    verify: "核验公共讨论、用户语言、文化符号和既有内容热度。",
  },
  {
    type: "必须项", name: "真需求", weight: 16,
    question: "读者是否真的痛、焦虑、想要，或能获得明确收获？",
    risk: "只有编辑觉得好，读者没有迫切需要。",
    description: "判断读者是否真的痛、焦虑、想要，或能获得明确收获。",
    high: "需求尖锐、真实、可被读者主动表达，并有行为或付费证据。",
    low: "只有编辑觉得好，读者没有迫切需要。",
    verify: "开展读者访谈，核验评论、搜索、付费或替代行为。",
  },
  {
    type: "必须项", name: "购买理由", weight: 16,
    question: "内容价值是否容易被翻译成一句下单理由？",
    risk: "读者觉得好，但不等于会买。",
    description: "判断内容价值能否变成让读者下单的一句话。",
    high: "能说清买回去解决什么、获得什么或改变什么。",
    low: "只能获得点赞或收藏，难以转化为购买。",
    verify: "测试一句话卖点、副标题、样章落地页和小额投放转化。",
  },
  {
    type: "必须项", name: "市场位置", weight: 16,
    question: "竞品强不强？本书是否有可挑战、可切入的位置？",
    risk: "找不到进入市场的角度，营销会很困难。",
    description: "判断竞品格局和本书的切入角度。",
    high: "有明确差异化，如形态、定价、场景、人群或叙事角度。",
    low: "竞品强且同质化，进入市场没有理由。",
    verify: "比较直接/间接竞品的读者、价格、卖点、渠道和评价。",
  },
  {
    type: "必须项", name: "纸质书匹配度", weight: 16,
    question: "这个需求为什么一定或更适合做成纸质书？",
    risk: "可能更适合短视频、课程或工具表格。",
    description: "判断需求为何适合做成纸质书，而非一次性内容。",
    high: "需要收藏、反复阅读、系统学习、礼赠、展示或长期陪伴。",
    low: "更适合短视频、课程、工具表格或即时搜索。",
    verify: "验证阅读深度、使用周期、礼赠场景、装帧价值和纸电差异。",
  },
  {
    type: "加分项", name: "公司资源匹配度", weight: 10,
    question: "现有达人、渠道、社群、私域和机构合作能否触达目标读者？",
    risk: "需要从零开路，周期长、成本高。",
    description: "判断现有资源是否能触达目标读者并形成验证闭环。",
    high: "已有资源或合作可直接触达并验证。",
    low: "需要从零开路，周期长、成本高。",
    verify: "列出资源、负责人、历史数据、触达规模和真实承诺。",
  },
  {
    type: "加分项", name: "长期/战略价值", weight: 10,
    question: "它能否成为长销品、系列品、品牌品或未来增长入口？",
    risk: "只能靠短期流量打一波，后续价值弱。",
    description: "判断能否成为长销品、系列品、品牌品或增长入口。",
    high: "能沉淀品牌、用户、产品线或后续业务。",
    low: "只能靠短期流量打一波，后续价值弱。",
    verify: "验证复购、系列延展、常青需求、作者成长性和品牌资产。",
  },
];

const byName = new Map((input.dimensions || []).map((item) => [item.name, item]));
const workbook = Workbook.create();
const summary = workbook.worksheets.add("评估总表");
const guide = workbook.worksheets.add("维度说明");
const rules = workbook.worksheets.add("评分规则");
workbook.comments.setSelf({ displayName: input.reviewer || "User" });

const colors = {
  navy: "#243447", blue: "#315B7D", paleBlue: "#EAF1F6", paper: "#F8F6F1",
  ink: "#1F2933", muted: "#66727E", line: "#CBD3DA", orange: "#D97735",
  paleOrange: "#FBE9DC", green: "#427A5B", paleGreen: "#E5F1E9", white: "#FFFFFF",
  input: "#FFF8DE",
};
const border = { preset: "all", style: "thin", color: colors.line };

summary.showGridLines = false;
summary.mergeCells("A1:L2");
summary.getRange("A1").values = [["出版选题七问评估表"]];
summary.getRange("A1:L2").format = {
  fill: colors.navy, font: { bold: true, color: colors.white, size: 20 },
  horizontalAlignment: "center", verticalAlignment: "center",
};
summary.getRange("A3:L3").values = [["书名/项目", input.title || "", "评分日期", input.date || new Date(), "评审人", input.reviewer || "", "总分", null, "致命短板检查", null, "建议结论", null]];
summary.getRange("A3:L3").format = { fill: colors.paper, font: { bold: true, color: colors.ink }, borders: border, verticalAlignment: "center" };
summary.getRange("B3,D3,F3").format = { fill: colors.input };
summary.getRange("D3").format.numberFormat = "yyyy-mm-dd";
summary.getRange("H3").formulas = [["=IF(OR(COUNT(D7:D11)<5,D12=\"\",D13=\"\"),\"\",SUM(E7:E13))"]];
summary.getRange("J3").formulas = [["=IF(COUNT(D7:D11)<5,\"待评分\",IF(COUNTIF(D7:D11,\"<2\")>0,\"有（必须项低于2分）\",\"无（五项均不低于2分）\"))"]];
summary.getRange("L3").formulas = [["=IF(COUNT(D7:D11)<5,\"待评分\",IF(OR(D12=\"\",D13=\"\"),\"待用户选择\",IF(COUNTIF(D7:D11,\"<2\")>0,\"暂缓/重做选题\",IF(H3>85,\"重点加码\",IF(H3>=70,\"可以推进\",IF(H3>=55,\"小预算验证\",\"不建议重点投入\"))))))"]];
summary.getRange("H3,J3,L3").format = { fill: colors.paleBlue, font: { bold: true, color: colors.navy }, horizontalAlignment: "center" };
summary.getRange("A4").values = [["一句话结论"]];
summary.mergeCells("B4:L4");
summary.getRange("B4").values = [[input.oneLineConclusion || "待生成：完成评分后，按“决策等级：最强优势、关键短板和建议动作”的顺序填写。"]];
summary.getRange("A4").format = { fill: colors.navy, font: { bold: true, color: colors.white }, verticalAlignment: "center" };
summary.getRange("B4:L4").format = { fill: colors.paper, font: { bold: true, color: colors.ink }, borders: border, wrapText: true, verticalAlignment: "center" };
summary.mergeCells("A5:L5");
summary.getRange("A5").values = [["前 5 项由 AI 评分；后 2 项必须由用户选择高/中/低。两项都选完后才输出最终结论。"]];
summary.getRange("A5:L5").format = { fill: colors.paleOrange, font: { italic: true, color: colors.orange }, horizontalAlignment: "left" };
summary.getRange("A6:L6").values = [["类型", "维度", "权重", "评分/用户选择", "加权得分", "核心判断问题", "证据与来源", "证据等级", "判断理由", "低分风险", "下一步动作", "置信度"]];
summary.getRange("A6:L6").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };

const rows = dimensions.map((dimension) => {
  const supplied = byName.get(dimension.name) || {};
  return [
    dimension.type, dimension.name, dimension.weight,
    dimension.type === "必须项" ? (Number.isFinite(supplied.score) ? supplied.score : null) : (supplied.userSelection || ""),
    null, dimension.question, supplied.evidence || "", supplied.evidenceLevel || "",
    supplied.reasoning || "", dimension.risk, supplied.action || dimension.verify, supplied.confidence || "",
  ];
});
summary.getRange("A7:L13").values = rows;
for (let row = 7; row <= 11; row += 1) {
  summary.getRange(`E${row}`).formulas = [[`=IF(D${row}=\"\",\"\",D${row}/5*C${row})`]];
}
for (let row = 12; row <= 13; row += 1) {
  summary.getRange(`E${row}`).formulas = [[`=IF(D${row}=\"\",\"\",IF(D${row}=\"高\",C${row},IF(D${row}=\"中\",C${row}*0.6,IF(D${row}=\"低\",C${row}*0.2,\"\"))))`]];
}
summary.getRange("A7:L13").format = { borders: border, verticalAlignment: "top", wrapText: true, font: { color: colors.ink, size: 10 } };
summary.getRange("D7:D13").format = { fill: colors.input, horizontalAlignment: "center", numberFormat: "0" };
summary.getRange("G7:I13,K7:L13").format.fill = colors.input;
summary.getRange("C7:E13").format.horizontalAlignment = "center";
summary.getRange("E7:E13").format.numberFormat = "0.0";
summary.getRange("D7:D11").dataValidation = { rule: { type: "whole", operator: "between", formula1: 0, formula2: 5 } };
summary.getRange("D12:D13").dataValidation = { rule: { type: "list", values: ["高", "中", "低"] } };
summary.getRange("H7:H13").dataValidation = { rule: { type: "list", values: ["A 直接证据", "B 间接证据", "C 合理推断", "D 未知"] } };
summary.getRange("L7:L13").dataValidation = { rule: { type: "list", values: ["高", "中", "低"] } };
summary.getRange("D7:D11").conditionalFormats.add("cellIs", { operator: "lessThan", formula: 2, format: { fill: colors.paleOrange, font: { bold: true, color: colors.orange } } });
summary.getRange("J3").conditionalFormats.add("containsText", { text: "有（", format: { fill: colors.paleOrange, font: { bold: true, color: colors.orange } } });
summary.getRange("L3").conditionalFormats.add("containsText", { text: "重点加码", format: { fill: colors.paleGreen, font: { bold: true, color: colors.green } } });
summary.freezePanes.freezeRows(6);

const widths = [10, 18, 8, 11, 12, 32, 34, 15, 30, 28, 32, 11];
widths.forEach((width, index) => { summary.getRangeByIndexes(0, index, 13, 1).format.columnWidth = width; });
summary.getRange("1:2").format.rowHeight = 28;
summary.getRange("3:3").format.rowHeight = 34;
summary.getRange("4:4").format.rowHeight = 48;
summary.getRange("5:6").format.rowHeight = 30;
summary.getRange("7:13").format.rowHeight = 78;

guide.showGridLines = false;
guide.mergeCells("A1:E2");
guide.getRange("A1").values = [["七个维度说明"]];
guide.getRange("A1:E2").format = { fill: colors.navy, font: { bold: true, color: colors.white, size: 18 }, horizontalAlignment: "center", verticalAlignment: "center" };
guide.getRange("A4:E4").values = [["维度", "简要说明", "高分表现", "低分警讯", "补证方向"]];
guide.getRange("A4:E4").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center" };
guide.getRange("A5:E11").values = dimensions.map((d) => [d.name, d.description, d.high, d.low, d.verify]);
guide.getRange("A5:E11").format = { borders: border, wrapText: true, verticalAlignment: "top", font: { color: colors.ink, size: 10 } };
guide.getRange("A5:A11").format = { fill: colors.paleBlue, font: { bold: true, color: colors.blue } };
[18, 35, 38, 34, 38].forEach((width, index) => { guide.getRangeByIndexes(0, index, 11, 1).format.columnWidth = width; });
guide.getRange("5:11").format.rowHeight = 72;
guide.freezePanes.freezeRows(4);

rules.showGridLines = false;
rules.mergeCells("A1:D2");
rules.getRange("A1").values = [["评分规则与来源"]];
rules.getRange("A1:D2").format = { fill: colors.navy, font: { bold: true, color: colors.white, size: 18 }, horizontalAlignment: "center", verticalAlignment: "center" };
rules.getRange("A4:D4").values = [["项目", "规则", "说明", "备注"]];
rules.getRange("A4:D4").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center" };
rules.getRange("A5:D10").values = [
  ["权重", "必须项 80 分；加分项 20 分", "必须项每项 16 分；加分项每项 10 分。", "前 5 项由 AI 评估；后 2 项由用户选择。"],
  ["致命短板", "任一必须项 <2 分", "结论直接提示“暂缓/重做选题”。", "两个加分项不能抵消致命短板。"],
  ["重点加码", "总分 >85 且无致命短板", "进入重点资源配置讨论。", "85 分归入“可以推进”。"],
  ["可以推进", "总分 70–85 且无致命短板", "补强短板后推进。", ""],
  ["小预算验证", "总分 55–69 且无致命短板", "先做最低成本验证。", ""],
  ["不建议重点投入", "总分 <55 且无致命短板", "建议放弃、重构或等待条件变化。", ""],
];
rules.getRange("A5:D10").format = { borders: border, wrapText: true, verticalAlignment: "top" };
rules.getRange("A12:C12").values = [["分值", "通用含义", "证据要求"]];
rules.getRange("A12:C12").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center" };
rules.getRange("A13:C18").values = [
  [0, "完全没有或已有直接反证", "逻辑不成立或有明确反证"],
  [1, "很弱，仅有模糊感觉", "关键链条断裂"],
  [2, "不足，风险显著", "有零散线索但不足以支撑立项"],
  [3, "基本成立", "至少有一项可复核证据"],
  [4, "较强", "多项 A/B 级证据相互印证"],
  [5, "非常强", "证据充分且已被现实结果验证"],
];
rules.getRange("A13:C18").format = { borders: border, wrapText: true };
rules.getRange("A20:C20").values = [["证据等级", "定义", "示例"]];
rules.getRange("A20:C20").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center" };
rules.getRange("A21:C24").values = [
  ["A 直接证据", "来自对象或真实行为", "书稿、访谈、成交、投放转化、已确认合作"],
  ["B 间接证据", "可靠但非本项目直接验证", "竞品数据、作者历史表现、行业报告"],
  ["C 合理推断", "从材料推导但尚未验证", "根据目录或评论推测"],
  ["D 未知", "缺少支撑材料", "无读者研究、竞品数据或资源确认"],
];
rules.getRange("A21:C24").format = { borders: border, wrapText: true };
rules.getRange("A26:C26").values = [["加分项用户选择", "折算分值", "单项加权得分"]];
rules.getRange("A26:C26").format = { fill: colors.blue, font: { bold: true, color: colors.white }, borders: border, horizontalAlignment: "center" };
rules.getRange("A27:C29").values = [["高", 5, 10], ["中", 3, 6], ["低", 1, 2]];
rules.getRange("A27:C29").format = { borders: border, horizontalAlignment: "center" };
rules.mergeCells("A31:D31");
rules.getRange("A31").values = [["两个加分项必须由用户分别选择；AI不得代选或默认选择“中”。"]];
rules.getRange("A31:D31").format = { fill: colors.paleOrange, font: { bold: true, color: colors.orange }, wrapText: true };
[20, 32, 42, 42].forEach((width, index) => { rules.getRangeByIndexes(0, index, 31, 1).format.columnWidth = width; });
rules.getRange("5:10").format.rowHeight = 42;
rules.getRange("13:18").format.rowHeight = 32;
rules.getRange("21:24").format.rowHeight = 38;
rules.getRange("31:31").format.rowHeight = 32;
rules.freezePanes.freezeRows(4);

await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

if (verifyDir) {
  await fs.mkdir(verifyDir, { recursive: true });
  for (const [sheetName, range] of [["评估总表", "A1:L13"], ["维度说明", "A1:E11"], ["评分规则", "A1:D31"]]) {
    const png = await workbook.render({ sheetName, range, scale: 1.5, format: "png" });
    await fs.writeFile(path.join(verifyDir, `${sheetName}.png`), new Uint8Array(await png.arrayBuffer()));
  }
  const checks = await workbook.inspect({ kind: "table", range: "评估总表!A1:L13", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 12 });
  const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
  console.log(checks.ndjson);
  console.log(errors.ndjson);
}
console.log(`Saved ${outputPath}`);
