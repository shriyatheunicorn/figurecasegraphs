const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, LevelFormat, Table, TableRow, TableCell, WidthType, ShadingType, PageBreak } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function hCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: "1E2761", type: ShadingType.CLEAR }, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 18, bold: true, color: "FFFFFF" })] })] });
}
function dCell(text, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? { fill: "F1F5F9", type: ShadingType.CLEAR } : undefined, margins: cellMargins,
    children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), font: opts.mono ? "Consolas" : "Arial", size: 18, color: opts.color || "1E293B", bold: opts.bold || false })] })] });
}

const c = [];
const title = (t) => new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: t, font: "Arial", size: 44, bold: true, color: "1E2761" })] });
const h2 = (t) => new Paragraph({ spacing: { before: 300, after: 120 }, children: [new TextRun({ text: t, font: "Arial", size: 30, bold: true, color: "1E2761" })] });
const h3 = (t) => new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: t, font: "Arial", size: 24, bold: true, color: "065A82" })] });
const body = (t) => new Paragraph({ spacing: { after: 120, line: 276 }, children: [new TextRun({ text: t, font: "Arial", size: 20, color: "1E293B" })] });
const divider = () => new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DEE2E6" } }, spacing: { after: 200 }, children: [] });
const spacer = () => new Paragraph({ spacing: { after: 80 }, children: [] });
const pageBreak = () => new Paragraph({ children: [new TextRun({ children: [new PageBreak()] })] });

// ========== TITLE ==========
c.push(new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "FIGURE CASE STUDY", font: "Arial", size: 22, bold: true, color: "0D9488" })] }));
c.push(title("Complete Working Calculations"));
c.push(body("Every number, grouped aggregation, and derived metric used in the analysis — with the raw output for full transparency."));
c.push(divider());

// ========== SECTION 1 ==========
c.push(h2("Section 1: Raw Data Structure"));
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680], rows: [
  new TableRow({ children: [hCell("Metric", 4680), hCell("Value", 4680)] }),
  new TableRow({ children: [dCell("Total rows in file", 4680), dCell("7,835", 4680)] }),
  new TableRow({ children: [dCell("Unique dataset names", 4680, {shade:true}), dCell("64", 4680, {shade:true})] }),
  new TableRow({ children: [dCell("Rows with explicit package/box counts", 4680), dCell("195 (excluded from main analysis)", 4680)] }),
  new TableRow({ children: [dCell("Main analysis rows", 4680, {shade:true}), dCell("7,640", 4680, {shade:true})] }),
  new TableRow({ children: [dCell("Date range (main)", 4680), dCell("2025-11-14 17:33:15 to 2025-11-26 15:35:04", 4680)] }),
  new TableRow({ children: [dCell("Unique dates (main)", 4680, {shade:true}), dCell("10", 4680, {shade:true})] }),
  new TableRow({ children: [dCell("Unique configs (main)", 4680), dCell("15", 4680)] }),
  new TableRow({ children: [dCell("Unique pilots", 4680, {shade:true}), dCell("5 (Pilot1–Pilot5)", 4680, {shade:true})] }),
]}));
c.push(spacer());

c.push(h3("Field Extraction Results"));
c.push(body("Dates (10): 2025-11-14, 11-17, 11-18, 11-19, 11-20, 11-21, 11-22, 11-24, 11-25, 11-26"));
c.push(body("Configs (15): c-100, c-101, c-103, c-105, c-106, c-110, c-111, c-128, c-129, c-142, c-143, c-144, c-146, c-188, c-198"));
c.push(body("Quality: HQ = 6,792 episodes, ABC = 848 episodes"));
c.push(body("No-BB flag: No = 6,040, Yes = 1,600"));
c.push(body("Speed flag: No = 7,483, Yes = 157"));
c.push(divider());

// ========== SECTION 2 ==========
c.push(h2("Section 2: Descriptive Statistics (All 7,640 Main Episodes)"));
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4000, 2680, 2680], rows: [
  new TableRow({ children: [hCell("Metric", 4000), hCell("Episode Length (sec)", 2680), hCell("Implied Cycle (÷10)", 2680)] }),
  new TableRow({ children: [dCell("Count", 4000), dCell("7,640", 2680, {align:AlignmentType.RIGHT}), dCell("7,640", 2680, {align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("Mean", 4000, {shade:true}), dCell("37.014", 2680, {shade:true, align:AlignmentType.RIGHT}), dCell("3.701", 2680, {shade:true, align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("Median", 4000), dCell("36.844", 2680, {align:AlignmentType.RIGHT}), dCell("3.684", 2680, {align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("Std Dev", 4000, {shade:true}), dCell("9.260", 2680, {shade:true, align:AlignmentType.RIGHT}), dCell("0.926", 2680, {shade:true, align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("Min", 4000), dCell("0.029", 2680, {align:AlignmentType.RIGHT}), dCell("0.003", 2680, {align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("Max", 4000, {shade:true}), dCell("115.955", 2680, {shade:true, align:AlignmentType.RIGHT}), dCell("11.596", 2680, {shade:true, align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("5th Percentile", 4000), dCell("23.805", 2680, {align:AlignmentType.RIGHT}), dCell("2.381", 2680, {align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("25th Percentile", 4000, {shade:true}), dCell("31.705", 2680, {shade:true, align:AlignmentType.RIGHT}), dCell("3.171", 2680, {shade:true, align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("75th Percentile", 4000), dCell("42.487", 2680, {align:AlignmentType.RIGHT}), dCell("4.249", 2680, {align:AlignmentType.RIGHT})] }),
  new TableRow({ children: [dCell("95th Percentile", 4000, {shade:true}), dCell("51.450", 2680, {shade:true, align:AlignmentType.RIGHT}), dCell("5.145", 2680, {shade:true, align:AlignmentType.RIGHT})] }),
]}));
c.push(spacer());
c.push(body("Target: 3.500s per package. Current mean is 5.8% above target (gap = 0.201s)."));
c.push(body("Episodes under 35s (target equivalent): 3,090 / 7,640 = 40.4%"));
c.push(body("Outliers < 5s: 44 episodes. Outliers > 80s: 6 episodes."));
c.push(pageBreak());

// ========== SECTION 3: OPERATOR ==========
c.push(h2("Section 3: Operator Analysis"));
const opRows = [
  ["Pilot3", "2,753", "35.57", "35.38", "9.15", "3.56", "48.0%", "Nov 17–26 (8 days)"],
  ["Pilot2", "1,317", "36.03", "35.72", "9.23", "3.60", "46.8%", "Nov 17–25 (8 days)"],
  ["Pilot4", "848", "37.42", "37.22", "8.88", "3.74", "36.4%", "Nov 14–25 (7 days)"],
  ["Pilot5", "1,336", "38.62", "38.77", "9.94", "3.86", "30.5%", "Nov 17–26 (8 days)"],
  ["Pilot1", "1,386", "39.01", "38.31", "8.40", "3.90", "31.4%", "Nov 17–25 (6 days)"],
];
const opColW = [900, 900, 1100, 1100, 900, 1000, 1000, 2460];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: opColW, rows: [
  new TableRow({ children: ["Pilot", "Episodes", "Mean (s)", "Median (s)", "Std (s)", "Cycle (s)", "% < 3.5s", "Active Dates"].map((h,i) => hCell(h, opColW[i])) }),
  ...opRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, opColW[i], { shade: idx%2===1, color: i===5 && parseFloat(val)<3.5 ? "0D9488" : "1E293B", bold: i===5 && parseFloat(val)<3.5 })) }))
]}));
c.push(spacer());
c.push(body("Fastest to slowest spread: 3.56s (Pilot3) to 3.90s (Pilot1) = 0.34s gap (9.6%)."));
c.push(divider());

// ========== SECTION 4: CONFIG ==========
c.push(h2("Section 4: Config Analysis"));
const cfgRows = [
  ["c-198", "565", "32.14", "31.64", "7.48", "3.21", "65.7%", "Nov 25–26"],
  ["c-144", "2,086", "34.07", "34.02", "8.20", "3.41", "56.0%", "Nov 20–24"],
  ["c-129", "11", "35.35", "39.78", "15.30", "3.53", "36.4%", "Nov 18"],
  ["c-188", "813", "35.35", "34.97", "8.38", "3.53", "50.3%", "Nov 21–25"],
  ["c-142", "326", "36.25", "36.54", "10.33", "3.63", "39.3%", "Nov 21"],
  ["c-101", "155", "37.25", "35.90", "8.78", "3.72", "43.9%", "Nov 19–20"],
  ["c-100", "378", "37.79", "37.21", "8.70", "3.78", "37.6%", "Nov 14–18"],
  ["c-110", "385", "38.56", "38.49", "8.53", "3.86", "30.9%", "Nov 17–20"],
  ["c-105", "197", "39.06", "38.74", "7.51", "3.91", "27.9%", "Nov 25"],
  ["c-143", "839", "39.17", "38.63", "8.37", "3.92", "29.2%", "Nov 18–25"],
  ["c-128", "13", "40.05", "39.72", "9.65", "4.01", "15.4%", "Nov 20"],
  ["c-111", "177", "40.15", "37.38", "12.49", "4.01", "36.7%", "Nov 17, 20"],
  ["c-146", "57", "40.43", "40.81", "9.31", "4.04", "24.6%", "Nov 25"],
  ["c-103", "1,312", "41.02", "41.68", "9.49", "4.10", "17.1%", "Nov 17–19"],
  ["c-106", "326", "41.13", "41.48", "9.72", "4.11", "22.7%", "Nov 20"],
];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: opColW, rows: [
  new TableRow({ children: ["Config", "Episodes", "Mean (s)", "Median (s)", "Std (s)", "Cycle (s)", "% < 3.5s", "Active Dates"].map((h,i) => hCell(h, opColW[i])) }),
  ...cfgRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, opColW[i], { shade: idx%2===1, color: i===5 && parseFloat(val)<3.5 ? "0D9488" : "1E293B", bold: i===5 && parseFloat(val)<3.5 })) }))
]}));
c.push(spacer());
c.push(body("Fastest to slowest: 3.21s (c-198) to 4.11s (c-106) = 0.90s gap (28%). Config variance is ~3x larger than operator variance."));
c.push(pageBreak());

// ========== SECTION 5: DAILY ==========
c.push(h2("Section 5: Daily Trends"));
const dayRows = [
  ["2025-11-14", "2", "54.97", "54.97", "5.50", "0.0%"],
  ["2025-11-17", "730", "40.77", "40.11", "4.08", "25.2%"],
  ["2025-11-18", "889", "39.28", "39.58", "3.93", "25.2%"],
  ["2025-11-19", "1,053", "39.65", "39.73", "3.97", "24.5%"],
  ["2025-11-20", "754", "39.68", "39.50", "3.97", "30.2%"],
  ["2025-11-21", "1,444", "34.39", "34.73", "3.44", "51.1%"],
  ["2025-11-22", "297", "29.80", "29.91", "2.98", "84.5%"],
  ["2025-11-24", "1,080", "36.91", "36.55", "3.69", "42.3%"],
  ["2025-11-25", "1,031", "35.81", "35.51", "3.58", "46.7%"],
  ["2025-11-26", "360", "30.59", "29.78", "3.06", "74.7%"],
];
const dayColW = [1600, 1100, 1500, 1500, 1500, 2160];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: dayColW, rows: [
  new TableRow({ children: ["Date", "Episodes", "Mean (s)", "Median (s)", "Cycle (s)", "% Under Target"].map((h,i) => hCell(h, dayColW[i])) }),
  ...dayRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, dayColW[i], { shade: idx%2===1, color: i===4 && parseFloat(val)<3.5 ? "0D9488" : "1E293B", bold: i===4 && parseFloat(val)<3.5 })) }))
]}));
c.push(spacer());
c.push(body("Nov 22 (2.98s) and Nov 26 (3.06s) are the only days with mean cycle time under target. Clear downward trend over the collection period."));
c.push(divider());

// ========== SECTION 6: SPEED ==========
c.push(h2("Section 6: Speed Mode Analysis"));
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3120, 3120, 3120], rows: [
  new TableRow({ children: [hCell("Metric", 3120), hCell("Speed Mode", 3120), hCell("Normal", 3120)] }),
  new TableRow({ children: [dCell("Episode Count", 3120), dCell("157", 3120), dCell("7,483", 3120)] }),
  new TableRow({ children: [dCell("Mean Ep Length (sec)", 3120, {shade:true}), dCell("28.74", 3120, {shade:true}), dCell("37.19", 3120, {shade:true})] }),
  new TableRow({ children: [dCell("Implied Cycle (sec)", 3120), dCell("2.87", 3120, {color:"0D9488", bold:true}), dCell("3.72", 3120)] }),
  new TableRow({ children: [dCell("Configs Used", 3120, {shade:true}), dCell("c-144, c-188", 3120, {shade:true}), dCell("All 15", 3120, {shade:true})] }),
]}));
c.push(spacer());
c.push(body("Speed mode is 22.7% faster (0.84s per package). Both speed-mode configs (c-144, c-188) are already among the better-performing configs."));
c.push(divider());

// ========== SECTION 7: NO-BB ==========
c.push(h2("Section 7: No Bounding Box Analysis"));
const bbRows = [
  ["c-101", "70", "4.14", "85", "3.38", "+0.76", "BB slower"],
  ["c-106", "182", "4.28", "144", "3.90", "+0.38", "BB slower"],
  ["c-110", "132", "3.90", "253", "3.83", "+0.07", "BB slower"],
  ["c-143", "622", "3.86", "217", "4.08", "-0.21", "No-BB slower"],
  ["c-144", "1,196", "3.38", "890", "3.44", "-0.06", "No-BB slower"],
];
const bbColW = [1000, 900, 1200, 1000, 1200, 1000, 3060];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: bbColW, rows: [
  new TableRow({ children: ["Config", "BB (n)", "BB Cycle", "No-BB (n)", "No-BB Cycle", "Diff", "Direction"].map((h,i) => hCell(h, bbColW[i])) }),
  ...bbRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, bbColW[i], { shade: idx%2===1 })) }))
]}));
c.push(spacer());
c.push(body("Impact is config-dependent. Removing bounding boxes helps on c-101 and c-106 but slightly hurts on c-143 and c-144. Not a clear universal improvement."));
c.push(divider());

// ========== SECTION 8: LEARNING ==========
c.push(h2("Section 8: Learning Curves (First 50 vs Last 50 Episodes)"));
const learnRows = [
  ["c-103", "1,312", "49.56", "37.90", "11.67", "23.5%"],
  ["c-144", "2,086", "43.94", "37.41", "6.53", "14.9%"],
  ["c-188", "813", "36.02", "33.41", "2.61", "7.3%"],
  ["c-198", "565", "36.74", "34.75", "2.00", "5.4%"],
  ["c-100", "378", "37.68", "35.92", "1.77", "4.7%"],
  ["c-143", "839", "36.77", "36.60", "0.16", "0.4%"],
  ["c-110", "385", "37.31", "39.31", "-2.00", "-5.4%"],
];
const learnColW = [1000, 1100, 1500, 1500, 1500, 2760];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: learnColW, rows: [
  new TableRow({ children: ["Config", "Total (n)", "First 50 (s)", "Last 50 (s)", "Improve (s)", "Improve (%)"].map((h,i) => hCell(h, learnColW[i])) }),
  ...learnRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, learnColW[i], { shade: idx%2===1 })) }))
]}));
c.push(spacer());
c.push(body("c-103 shows the largest learning effect (23.5% improvement). c-110 actually got worse, possibly due to confounding factors. Configs with fewer total episodes show less learning, as expected."));
c.push(pageBreak());

// ========== SECTION 9: CROSS TAB ==========
c.push(h2("Section 9: Operator × Config Cross-Tabulation"));
c.push(body("All 32 operator-config pairings, sorted by pilot then cycle time. Checkmark (✓) indicates pairings that beat the 3.5s target."));
c.push(spacer());

const oxRows = [
  ["Pilot1","c-111","128","35.9","3.59","11/20 00:16","11/20 02:59"],
  ["Pilot1","c-101","124","36.9","3.69","11/20 19:49","11/20 22:53"],
  ["Pilot1","c-110","335","38.8","3.88","11/17 17:20","11/19 23:34"],
  ["Pilot1","c-105","197","39.1","3.91","11/25 17:40","11/26 00:11"],
  ["Pilot1","c-143","564","39.6","3.96","11/18 17:34","11/24 22:25"],
  ["Pilot1","c-146","15","45.7","4.57","11/25 16:04","11/25 16:40"],
  ["Pilot1","c-100","23","52.9","5.29","11/17 15:11","11/17 17:14"],
  ["Pilot2","c-144","508","30.6","3.06 ✓","11/21 21:30","11/24 21:29"],
  ["Pilot2","c-101","22","36.8","3.68","11/19 23:38","11/20 00:16"],
  ["Pilot2","c-188","271","36.9","3.69","11/21 20:08","11/25 19:53"],
  ["Pilot2","c-106","146","38.7","3.87","11/20 16:56","11/20 21:00"],
  ["Pilot2","c-128","13","40.1","4.01","11/20 15:30","11/20 16:17"],
  ["Pilot2","c-103","357","41.8","4.18","11/17 23:59","11/19 22:47"],
  ["Pilot3","c-198","301","29.7","2.97 ✓","11/26 07:14","11/26 14:23"],
  ["Pilot3","c-188","350","33.0","3.30 ✓","11/25 07:38","11/25 14:54"],
  ["Pilot3","c-144","1321","34.3","3.43 ✓","11/21 07:40","11/24 14:56"],
  ["Pilot3","c-110","36","38.5","3.85","11/20 14:04","11/20 14:35"],
  ["Pilot3","c-103","555","40.0","4.00","11/18 07:13","11/19 14:39"],
  ["Pilot3","c-101","9","43.1","4.31","11/20 07:21","11/20 08:19"],
  ["Pilot3","c-106","132","43.4","4.34","11/20 09:21","11/20 11:59"],
  ["Pilot3","c-111","49","51.3","5.13","11/17 08:45","11/17 10:53"],
  ["Pilot4","c-110","14","33.2","3.32 ✓","11/19 08:20","11/19 08:48"],
  ["Pilot4","c-129","11","35.3","3.53","11/18 07:44","11/18 08:18"],
  ["Pilot4","c-100","355","36.8","3.68","11/14 17:33","11/18 14:26"],
  ["Pilot4","c-188","151","37.3","3.73","11/24 07:23","11/24 11:44"],
  ["Pilot4","c-143","275","38.4","3.84","11/19 09:22","11/25 10:26"],
  ["Pilot4","c-146","42","38.6","3.86","11/25 11:14","11/25 15:35"],
  ["Pilot5","c-198","264","35.0","3.50 ✓","11/25 20:33","11/26 15:35"],
  ["Pilot5","c-142","326","36.3","3.63","11/21 11:49","11/21 16:58"],
  ["Pilot5","c-188","41","37.7","3.77","11/21 18:28","11/21 20:01"],
  ["Pilot5","c-144","257","40.0","4.00","11/20 23:10","11/24 19:45"],
  ["Pilot5","c-103","400","41.7","4.17","11/17 18:40","11/19 19:26"],
  ["Pilot5","c-106","48","42.5","4.25","11/20 21:30","11/20 22:30"],
];
const oxColW = [900, 900, 900, 1200, 1200, 2130, 2130];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: oxColW, rows: [
  new TableRow({ children: ["Pilot","Config","n","Mean (s)","Cycle (s)","First","Last"].map((h,i) => hCell(h, oxColW[i])) }),
  ...oxRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, oxColW[i], {
    shade: idx%2===1,
    color: (i===4 && val.includes("✓")) ? "0D9488" : "1E293B",
    bold: (i===4 && val.includes("✓"))
  })) }))
]}));
c.push(spacer());
c.push(body("6 of 32 pairings beat the 3.5s target. Pilot3 accounts for 3 of them (c-198, c-188, c-144). Notable: Pilot2 on c-144 = 3.06s vs Pilot5 on c-144 = 4.00s — same config, 0.94s gap."));
c.push(pageBreak());

// ========== SECTION 10: TIME OF DAY ==========
c.push(h2("Section 10: Time-of-Day Analysis"));
const todRows = [
  ["00:00","107","40.0","4.00"], ["01:00","159","36.7","3.67"], ["02:00","50","32.3","3.23"],
  ["07:00","201","37.3","3.73"], ["08:00","448","37.6","3.76"], ["09:00","383","37.9","3.79"],
  ["10:00","520","36.3","3.63"], ["11:00","521","36.0","3.60"], ["12:00","402","34.7","3.47"],
  ["13:00","508","34.8","3.48"], ["14:00","671","35.0","3.50"], ["15:00","436","39.1","3.91"],
  ["16:00","485","37.5","3.75"], ["17:00","401","37.8","3.78"], ["18:00","422","37.4","3.74"],
  ["19:00","194","37.4","3.74"], ["20:00","460","40.4","4.04"], ["21:00","437","38.6","3.86"],
  ["22:00","485","36.1","3.61"], ["23:00","350","37.2","3.72"],
];
const todColW = [2000, 2000, 2680, 2680];
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: todColW, rows: [
  new TableRow({ children: ["Hour","Episodes","Mean Ep Length (s)","Implied Cycle (s)"].map((h,i) => hCell(h, todColW[i])) }),
  ...todRows.map((row, idx) => new TableRow({ children: row.map((val, i) => dCell(val, todColW[i], {
    shade: idx%2===1,
    color: (i===3 && parseFloat(val)<3.5) ? "0D9488" : "1E293B",
    bold: (i===3 && parseFloat(val)<3.5)
  })) }))
]}));
c.push(spacer());
c.push(body("Fastest hours: 12:00–14:00 (3.47–3.50s, near/at target). Slowest: 20:00 (4.04s) and 00:00 (4.00s). Midday performance is ~15% better than late evening, suggesting fatigue or shift-change effects."));
c.push(divider());

// ========== SECTION 11: QUALITY ==========
c.push(h2("Section 11: Quality Flag Analysis"));
c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3120, 3120, 3120], rows: [
  new TableRow({ children: [hCell("Metric", 3120), hCell("HQ", 3120), hCell("ABC", 3120)] }),
  new TableRow({ children: [dCell("Episodes", 3120), dCell("6,792", 3120), dCell("848", 3120)] }),
  new TableRow({ children: [dCell("Mean Ep Length (sec)", 3120, {shade:true}), dCell("36.96", 3120, {shade:true}), dCell("37.42", 3120, {shade:true})] }),
  new TableRow({ children: [dCell("Implied Cycle (sec)", 3120), dCell("3.70", 3120), dCell("3.74", 3120)] }),
  new TableRow({ children: [dCell("Std Dev (sec)", 3120, {shade:true}), dCell("9.31", 3120, {shade:true}), dCell("8.88", 3120, {shade:true})] }),
]}));
c.push(spacer());
c.push(body("Minimal difference between HQ and ABC quality episodes (0.04s). The quality flag does not appear to be a major driver of cycle time."));

const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } }, children: c }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/working_calculations.docx", buffer);
  console.log("Working calculations doc saved!");
});
