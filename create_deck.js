const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Operations Strategy Analyst";
pres.title = "Figure Case Study – Logistics Data Analysis & Ops Plan";

// Color palette - Midnight Executive
const C = {
  navy: "1E2761",
  ice: "CADCFC",
  white: "FFFFFF",
  dark: "0F172A",
  accent: "0D9488",
  secondary: "065A82",
  lightBg: "F8FAFB",
  text: "1E293B",
  muted: "64748B",
  red: "E74C3C",
  green: "10B981",
};

const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.12, angle: 135 });

// ============================================================
// SLIDE 1: Title Slide
// ============================================================
let slide1 = pres.addSlide();
slide1.background = { color: C.navy };

// Top accent bar
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

slide1.addText("FIGURE CASE STUDY", {
  x: 0.8, y: 1.0, w: 8.4, h: 0.6,
  fontSize: 16, fontFace: "Arial", color: C.accent, charSpacing: 6, bold: true, align: "left", margin: 0
});

slide1.addText("Logistics Data Analysis\n& Operational Plan", {
  x: 0.8, y: 1.7, w: 8.4, h: 2.0,
  fontSize: 40, fontFace: "Georgia", color: C.white, bold: true, align: "left", margin: 0,
  lineSpacingMultiple: 1.1
});

slide1.addText("Operations Strategy Analyst", {
  x: 0.8, y: 4.0, w: 8.4, h: 0.5,
  fontSize: 16, fontFace: "Calibri", color: C.ice, align: "left", margin: 0
});

// Bottom bar
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.secondary } });
slide1.addText("Prompt 1: Data Analysis  |  Prompt 2: Ops Plan", {
  x: 0.8, y: 5.2, w: 8.4, h: 0.425,
  fontSize: 12, fontFace: "Calibri", color: C.white, align: "left", margin: 0
});

slide1.addNotes(`[~2 minutes — Introduction]

Good morning/afternoon everyone. Thank you for taking the time to review this case study with me.

Today I'll walk you through two parts. First, a data analysis of our logistics teleoperation dataset — over 7,800 episodes of a robot placing packages onto a conveyor belt. I'll share key patterns, performance drivers, and recommendations for improving the autonomous policy we're training from this data.

Second, I'll present an operational plan for a 2-week data collection effort in a residential home, where we'll be teaching a robot to tidy up a playroom.

Let's start with the data analysis.`);

// ============================================================
// SLIDE 2: Methodology Overview
// ============================================================
let slideMeth = pres.addSlide();
slideMeth.background = { color: C.lightBg };

slideMeth.addText("Methodology", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Left column: Data parsing steps
slideMeth.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.4, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
slideMeth.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.4, h: 0.4, fill: { color: C.navy } });
slideMeth.addText("DATA EXTRACTION & PARSING", {
  x: 0.7, y: 1.12, w: 4.0, h: 0.4,
  fontSize: 11, fontFace: "Arial", color: C.white, bold: true, margin: 0
});

const parseSteps = [
  { step: "1. Load raw data", detail: "7,835 rows × 3 columns (Dataset Name, Episode ID, Episode Length). Skipped blank header row, dropped empty Column A." },
  { step: "2. Parse Dataset Name", detail: "Reverse-engineered the naming convention to extract: Date (first 8 chars), Config (regex: c-\\d+), Pilot (regex: Pilot\\d+), Quality (hq/abc), No-BB flag, Speed flag." },
  { step: "3. Extract timestamps", detail: "Parsed first 14 chars of Episode ID as YYYYMMDDHHmmss to get exact timestamps for each episode." },
  { step: "4. Filter dataset", detail: "Excluded 195 early pilot episodes with explicit package counts (different task setup). Main analysis: 7,640 episodes." },
  { step: "5. Key assumption", detail: "Divided episode length by 10 to get implied cycle time. Based on: \"10_packages\" label in explicit episodes, mean ~37s matching 10 × 3.5s target." },
];

parseSteps.forEach((item, i) => {
  const yPos = 1.6 + i * 0.7;
  slideMeth.addText(item.step, { x: 0.7, y: yPos, w: 4.0, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  slideMeth.addText(item.detail, { x: 0.7, y: yPos + 0.22, w: 4.0, h: 0.42, fontSize: 9, fontFace: "Calibri", color: C.muted, margin: 0 });
});

// Right column: Calculations performed
slideMeth.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.4, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
slideMeth.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.4, h: 0.4, fill: { color: C.secondary } });
slideMeth.addText("CALCULATIONS & ANALYSIS", {
  x: 5.3, y: 1.12, w: 4.0, h: 0.4,
  fontSize: 11, fontFace: "Arial", color: C.white, bold: true, margin: 0
});

const calcSteps = [
  { step: "Descriptive Statistics", detail: "Mean, median, std dev, min/max, percentiles (5th/95th) on episode length. All via live Excel formulas (AVERAGE, MEDIAN, STDEV, PERCENTILE)." },
  { step: "Target Comparison", detail: "COUNTIF < 35s / total count → 41% of episodes already under 3.5s cycle time target." },
  { step: "Grouped Aggregations", detail: "Grouped by Operator (5 pilots), Config (15 configs), Date (10 days), and Operator×Config (32 pairings). Computed count, mean, median, std, % under target for each." },
  { step: "Speed Mode & No-BB", detail: "AVERAGEIF with wildcard on dataset name to isolate speed (n=157) and no-BB (n=1,600) subsets. Compared against normal baselines." },
  { step: "Learning Curves", detail: "Sorted episodes chronologically within each config, compared first 50 vs. last 50 episode means to quantify ramp-up effects." },
];

calcSteps.forEach((item, i) => {
  const yPos = 1.6 + i * 0.7;
  slideMeth.addText(item.step, { x: 5.3, y: yPos, w: 4.0, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.secondary, bold: true, margin: 0 });
  slideMeth.addText(item.detail, { x: 5.3, y: yPos + 0.22, w: 4.0, h: 0.42, fontSize: 9, fontFace: "Calibri", color: C.muted, margin: 0 });
});

slideMeth.addNotes(`[~3 minutes — Methodology]

Before diving into findings, let me quickly walk through how I approached the data.

The raw dataset has 7,835 rows with three columns: a dataset name string, an episode ID, and the episode length in seconds. The dataset name is the richest field — it's a structured string that encodes the collection date, robot configuration, operator identity, and several flags like quality level and speed mode. There's no data dictionary, so I reverse-engineered the naming convention by examining all 64 unique dataset names.

I used regex pattern matching to extract each field. For example, robot configs follow a "c-" plus digits pattern, and operators are labeled "Pilot1" through "Pilot5." The episode ID also contains an embedded timestamp in YYYYMMDDHHmmss format, which I parsed to get exact timing for each episode.

One important filtering decision: 195 early episodes explicitly specify package counts in their names — like "1 package 1 box" or "10 packages." These represent different task setups, so I excluded them from the main analysis to keep comparisons clean.

The biggest assumption is dividing episode length by 10 to get per-package cycle time. This is supported by the "10 packages" label in explicit episodes, and by the fact that the overall mean of 37 seconds closely matches 10 packages at 3.5 seconds each. But it remains an assumption — I flag this throughout the analysis.

On the calculation side, all summary statistics use live Excel formulas so the spreadsheet stays dynamic. The grouped analyses — by operator, config, date, and operator-config pairs — were done in Python using pandas, since parsing the dataset names requires regex which Excel can't handle cleanly. Results are written into dedicated spreadsheet tabs.

Full methodology details are in the separate methodology document I've prepared.`);

// ============================================================
// SLIDE 3: Section Divider – Prompt 1
// ============================================================
let slide2 = pres.addSlide();
slide2.background = { color: C.lightBg };

slide2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.accent } });

slide2.addText("PROMPT 1", {
  x: 0.8, y: 1.4, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Arial", color: C.accent, charSpacing: 4, bold: true, margin: 0
});
slide2.addText("Data Analysis", {
  x: 0.8, y: 1.9, w: 8.4, h: 1.0,
  fontSize: 38, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});
slide2.addText("7,640 main episodes  •  15 robot configurations  •  5 operators  •  10 collection days", {
  x: 0.8, y: 3.1, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Calibri", color: C.muted, margin: 0
});

// Key stats row
const stats = [
  { val: "3.70s", label: "Avg Cycle Time", color: C.red },
  { val: "3.5s", label: "Target", color: C.accent },
  { val: "41%", label: "Under Target", color: C.secondary },
  { val: "2.87s", label: "Best (Speed Mode)", color: C.green },
];
stats.forEach((s, i) => {
  const xPos = 0.8 + i * 2.2;
  slide2.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.0, w: 2.0, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  slide2.addText(s.val, { x: xPos, y: 4.05, w: 2.0, h: 0.7, fontSize: 26, fontFace: "Georgia", color: s.color, bold: true, align: "center", margin: 0 });
  slide2.addText(s.label, { x: xPos, y: 4.7, w: 2.0, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.muted, align: "center", margin: 0 });
});

slide2.addNotes(`[~4 minutes — Dataset Overview & Key Metrics]

Let me orient you on the dataset. We collected 7,835 teleoperation episodes across 10 collection days in November 2025. After filtering out 195 early pilot episodes with different task setups, our main analysis covers 7,640 episodes spanning 15 robot configurations — labeled c-100 through c-198 — and 5 teleoperators, Pilot 1 through 5.

Each episode represents a full run of placing packages onto the conveyor belt with barcodes facing down. The episode length captures the total time, and since most episodes involve 10 packages, I've derived an "implied cycle time" by dividing episode length by 10.

Looking at the headline numbers: our current average cycle time is 3.70 seconds per package, which is about 6% above the 3.5-second target. However — and this is important — 41% of episodes already meet the target. And when operators use "speed mode," the average drops to 2.87 seconds, well below target. So the 3.5-second goal is clearly achievable — the question is what's dragging the average up and how we filter or improve the training data accordingly.

Let me dig into the trends.`);

// ============================================================
// SLIDE 3: Cycle Time Timeline + Distribution
// ============================================================
let slide3 = pres.addSlide();
slide3.background = { color: C.lightBg };

slide3.addText("Cycle Time Trends & Distribution", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Left chart: timeline
slide3.addImage({ path: path.resolve("/home/claude/chart1_timeline.png"), x: 0.3, y: 1.0, w: 5.4, h: 2.7 });

// Right chart: distribution
slide3.addImage({ path: path.resolve("/home/claude/chart4_distribution.png"), x: 5.3, y: 1.0, w: 4.5, h: 2.7 });

// Insights box
slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 9.0, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 0.08, h: 1.5, fill: { color: C.accent } });

slide3.addText("KEY FINDINGS", {
  x: 0.8, y: 3.9, w: 8.5, h: 0.35,
  fontSize: 11, fontFace: "Arial", color: C.accent, bold: true, charSpacing: 3, margin: 0
});
slide3.addText([
  { text: "Cycle times improved significantly over the collection period — Nov 22 (2.98s) and Nov 26 (3.06s) met the 3.5s target", options: { bullet: true, breakLine: true, fontSize: 11, color: C.text } },
  { text: "Overall mean (3.70s) is 6% above target; however, the latest days trend below it, suggesting the policy can reach 3.5s", options: { bullet: true, breakLine: true, fontSize: 11, color: C.text } },
  { text: "Distribution is right-skewed — 41% of episodes already under target, with a long tail of slow outliers worth filtering", options: { bullet: true, fontSize: 11, color: C.text } },
], { x: 0.8, y: 4.25, w: 8.5, h: 1.0, fontFace: "Calibri", margin: 0, paraSpaceAfter: 4 });

slide3.addNotes(`[~5 minutes — Cycle Time Trends & Distribution]

On the left, you can see cycle time plotted over the collection period. The trend is clearly downward — we started around 5.5 seconds in mid-October, and by the last two collection days — November 22nd and 26th — we're averaging under 3 seconds. That's a dramatic improvement.

A few things are driving this. First, there's a learning curve — operators get faster with practice. Second, we introduced newer robot configurations over time, and those newer configs tend to perform better. I'll break that down on the next slide.

On the right is the distribution of all cycle times. You can see the bulk of episodes cluster between 2.5 and 5 seconds, with a peak around 3.5 to 4 seconds. The red dashed line is our target. 41% of episodes — everything to the left of that line — already meet it.

There are two tails worth noting. On the left, we have 44 episodes under 0.5 seconds — these are almost certainly aborted or erroneously recorded episodes, not real completions. On the right, we have a long tail extending past 8 seconds, which likely represents episodes with errors, pauses, or difficult package orientations.

Both tails represent data quality issues. My recommendation is to filter these out before training — they'd teach the policy either incomplete trajectories or inefficient ones. I'd suggest a window of roughly 2 to 6 seconds per package as a clean training range, though that threshold should be validated with the robotics team.`);

// ============================================================
// SLIDE 4: Operator & Config Performance
// ============================================================
let slide4 = pres.addSlide();
slide4.background = { color: C.lightBg };

slide4.addText("Performance by Operator & Configuration", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Left chart: operators
slide4.addImage({ path: path.resolve("/home/claude/chart2_operators.png"), x: 0.2, y: 1.0, w: 4.9, h: 2.6 });

// Right chart: configs
slide4.addImage({ path: path.resolve("/home/claude/chart3_configs.png"), x: 5.0, y: 1.0, w: 4.9, h: 2.6 });

// Insights
slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9.0, h: 1.55, fill: { color: C.white }, shadow: makeShadow() });
slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 0.08, h: 1.55, fill: { color: C.accent } });

slide4.addText("KEY FINDINGS", {
  x: 0.8, y: 3.85, w: 8.5, h: 0.35,
  fontSize: 11, fontFace: "Arial", color: C.accent, bold: true, charSpacing: 3, margin: 0
});
slide4.addText([
  { text: "Operator variance: Pilot2 and Pilot3 are fastest (3.55s, 3.58s) — Pilot5 is slowest (3.86s). A ~9% gap suggests coaching or technique sharing could lift slower operators.", options: { bullet: true, breakLine: true, fontSize: 11, color: C.text } },
  { text: "Config matters: c-198 (3.21s) and c-144 (3.41s) beat target; c-103 (4.01s) and c-106 (4.11s) lag significantly. Newer configs outperform older ones.", options: { bullet: true, breakLine: true, fontSize: 11, color: C.text } },
  { text: "Learning curve: c-103 improved 11.7s (first 50 vs. last 50 eps), c-144 improved 6.5s. Early episodes inflate averages — consider excluding ramp-up data for training.", options: { bullet: true, fontSize: 11, color: C.text } },
], { x: 0.8, y: 4.2, w: 8.5, h: 1.1, fontFace: "Calibri", margin: 0, paraSpaceAfter: 4 });

slide4.addNotes(`[~6 minutes — Operator & Configuration Performance]

This is where it gets interesting. On the left, we have operator performance. Pilot 3 and Pilot 2 are our fastest operators at 3.56 and 3.60 seconds respectively — both very close to target. Pilot 5 and Pilot 1 are the slowest at 3.86 and 3.90 seconds. That's roughly a 9% gap between best and worst.

Now, 9% might not sound like much, but when you're trying to close a 6% gap to target, operator technique matters. I'd recommend having Pilot 2 and 3 share their approach with the slower operators — there may be specific teleoperation techniques or strategies they use that others could adopt.

On the right is the more impactful finding: robot configuration variance. This is a much bigger driver than operator skill. Config c-198, our newest, averages 3.21 seconds — well under target. Config c-144, our largest dataset with over 2,000 episodes, hits 3.41 seconds. But older configs like c-103 and c-106 are stuck above 4 seconds.

There's an important caveat here: configs were introduced sequentially, not tested in parallel. So the improvement from c-103 to c-198 could partly reflect operator learning rather than purely hardware or software changes. We'd need controlled testing — same operator, same day, different configs — to fully separate these effects.

I also looked at learning curves within individual configs. Config c-103 improved by 11.7 seconds comparing its first 50 episodes to its last 50. Config c-144 improved by 6.5 seconds. This suggests early episodes in any config are significantly slower as operators ramp up. For training data, we may want to exclude the first N episodes per config to avoid teaching the policy suboptimal early behavior.

One more data point from the cross-tabulation: Pilot 2 on config c-144 averages 3.06 seconds, while Pilot 5 on the same config averages 4.00 seconds. Same hardware, nearly 1 second gap — that confirms operator skill and config interact in ways we should understand better.`);

// ============================================================
// SLIDE 5: Speed Mode, No-BB, & Recommendations
// ============================================================
let slide5 = pres.addSlide();
slide5.background = { color: C.lightBg };

slide5.addText("Insights & Next Steps", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Speed mode chart
slide5.addImage({ path: path.resolve("/home/claude/chart5_speed.png"), x: 0.3, y: 1.0, w: 4.0, h: 2.3 });

// Insight cards on the right
const cards = [
  { title: "Speed Mode = 23% Faster", body: "Speed-mode episodes average 2.87s vs. 3.72s normal. This confirms the 3.5s target is achievable with optimized teleoperation technique.", color: C.green },
  { title: "'no_bbera' Flag: Unknown Variable", body: "~1,600 episodes carry this flag. Performance impact is mixed across configs (+2s on some, -2s on others). Flag meaning should be clarified before drawing conclusions.", color: C.secondary },
  { title: "44 Outliers <5s", body: "Likely incomplete/aborted episodes. These should be flagged and excluded from training data to avoid teaching the policy bad trajectories.", color: C.red },
];

cards.forEach((c, i) => {
  const yPos = 1.0 + i * 0.85;
  slide5.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: yPos, w: 5.0, h: 0.75, fill: { color: C.white }, shadow: makeShadow() });
  slide5.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: yPos, w: 0.07, h: 0.75, fill: { color: c.color } });
  slide5.addText(c.title, { x: 4.85, y: yPos + 0.02, w: 4.6, h: 0.28, fontSize: 12, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  slide5.addText(c.body, { x: 4.85, y: yPos + 0.30, w: 4.6, h: 0.42, fontSize: 10, fontFace: "Calibri", color: C.muted, margin: 0 });
});

// Recommended analyses box at bottom
slide5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.45, w: 9.0, h: 2.05, fill: { color: C.navy } });
slide5.addText("RECOMMENDED ADDITIONAL ANALYSES", {
  x: 0.8, y: 3.48, w: 8.5, h: 0.25,
  fontSize: 9, fontFace: "Arial", color: C.accent, bold: true, charSpacing: 3, margin: 0
});
slide5.addText([
  { text: "Per-package breakdown: Decompose episode time into individual pick-place cycles to identify slowest positions", options: { bullet: true, breakLine: true, fontSize: 9.5, color: C.ice } },
  { text: "Failure mode analysis: Correlate episode length with barcode scan pass rate to find quality-speed tradeoffs", options: { bullet: true, breakLine: true, fontSize: 9.5, color: C.ice } },
  { text: "Time-of-day effects: Noon episodes ~3s faster than 8 PM — fatigue may degrade quality late in shifts", options: { bullet: true, breakLine: true, fontSize: 9.5, color: C.ice } },
  { text: "Controlled pairing: Run best operators on best configs (c-198/c-144) to establish performance ceiling", options: { bullet: true, breakLine: true, fontSize: 9.5, color: C.ice } },
  { text: "Data filtering: Exclude ramp-up episodes, outliers (<5s, >80s), and non-HQ data for training quality", options: { bullet: true, fontSize: 9.5, color: C.ice } },
], { x: 0.8, y: 3.8, w: 8.5, h: 1.6, fontFace: "Calibri", margin: 0, paraSpaceAfter: 2 });

slide5.addNotes(`[~6 minutes — Key Insights & Recommended Next Steps]

Let me highlight three specific findings before moving to recommendations.

First, speed mode. About 157 episodes were collected in an explicit "speed mode," and they average 2.87 seconds — 23% faster than normal operation. This is our strongest evidence that the 3.5-second target is not only achievable but beatable. The question for the team is whether speed-mode data should be weighted more heavily in training, or whether it introduces behaviors that sacrifice reliability — like missed barcode orientations.

Second, the "no_bbera" flag. About 1,600 episodes carry this flag in their dataset name, but I don't know what "bbera" refers to — I initially assumed bounding boxes, but that turned out to be incorrect. What I can say is that the flag correlates with measurable performance differences: on some configs it's associated with faster episodes, on others slower. Before we draw any conclusions, we need to clarify with the team what this variable actually controls. Once we know that, a controlled A/B test would tell us whether to standardize it.

Third, outliers. We have 44 episodes under 5 seconds and 6 over 80 seconds. The short ones are almost certainly incomplete — you physically cannot place 10 packages in under 5 seconds. The long ones likely involve robot errors, operator breaks, or connectivity issues. Both should be flagged and excluded from training data.

Now for recommended additional analyses — and I want to emphasize these are things I think would meaningfully improve policy quality, not just academic exercises:

The most valuable would be a per-package breakdown. Right now, we only have total episode time. If we can decompose each episode into individual pick-place cycles, we could identify which package positions or orientations are systematically slower. That would tell us where the policy needs the most training signal.

Second, correlating episode length with scan success rate. We're optimizing for speed, but the constraint is barcode-down placement. Understanding the quality-speed frontier would help us set the right training target — maybe 3.5 seconds is too aggressive if it leads to scan failures, or maybe we can push even faster.

Third, time-of-day effects. I noticed episodes around noon are about 3 seconds faster than episodes at 8 PM. This could reflect operator fatigue, lighting changes, or shift patterns. If late-shift data is systematically worse, we should either downweight it or adjust scheduling.

Fourth, controlled pairings. Run our best operators on our best configs — Pilot 2 and 3 on c-198 and c-144 — to establish a true performance ceiling. That tells us the upper bound of what's possible with current hardware.

And fifth, a data filtering strategy. Combining everything — exclude ramp-up episodes, outliers, non-HQ quality data — to create a curated training set. The current dataset has a lot of noise, and a cleaner subset might train a better policy even if it's smaller.`);

// ============================================================
// SLIDE 6: Section Divider – Prompt 2
// ============================================================
let slide6 = pres.addSlide();
slide6.background = { color: C.lightBg };

slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.accent } });

slide6.addText("PROMPT 2", {
  x: 0.8, y: 1.4, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Arial", color: C.accent, charSpacing: 4, bold: true, margin: 0
});
slide6.addText("Operational Plan", {
  x: 0.8, y: 1.9, w: 8.4, h: 1.0,
  fontSize: 38, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});
slide6.addText("2-Week Playroom Tidying Data Collection — Scene Setup, Inventory & Variation Strategy", {
  x: 0.8, y: 3.1, w: 8.4, h: 0.5,
  fontSize: 14, fontFace: "Calibri", color: C.muted, margin: 0
});

// Summary metrics
const opStats = [
  { val: "10 Days", label: "Collection Window" },
  { val: "1 Robot", label: "8-hr Shifts" },
  { val: "~80 hrs", label: "Total Runtime" },
  { val: "150+", label: "Unique Objects" },
];
opStats.forEach((s, i) => {
  const xPos = 0.8 + i * 2.2;
  slide6.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.0, w: 2.0, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  slide6.addText(s.val, { x: xPos, y: 4.05, w: 2.0, h: 0.7, fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, align: "center", margin: 0 });
  slide6.addText(s.label, { x: xPos, y: 4.7, w: 2.0, h: 0.4, fontSize: 11, fontFace: "Calibri", color: C.muted, align: "center", margin: 0 });
});

slide6.addNotes(`[~3 minutes — Transition to Ops Plan]

Now let's shift to the second part: the operational plan for our 2-week playroom data collection.

The setup: we have one robot, one residential home about an hour from the office, and 8-hour shifts each day. The task is tidying up a playroom — picking up toys and putting them away. The critical requirement is dataset diversity — we need a wide variety of objects, placements, and tidying approaches to train a robust policy.

Let me walk you through how I'm thinking about this across three areas: what we need to buy, how we vary the scenes, and the day-by-day schedule.

A few key numbers to keep in mind: 10 working days across the 2 weeks, roughly 80 hours of total robot runtime, and my target is to collect data on at least 150 unique objects. The weekend in between gives us a natural checkpoint to review data quality and adjust the plan for week 2 — something I learned from the logistics dataset, where we saw performance change significantly over time.`);

// ============================================================
// SLIDE 7: Inventory & Scene Variation Plan
// ============================================================
let slide7 = pres.addSlide();
slide7.background = { color: C.lightBg };

slide7.addText("Inventory Acquisition & Scene Variation", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Left column: Inventory categories
slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.4, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.4, h: 0.4, fill: { color: C.navy } });
slide7.addText("OBJECT INVENTORY (150+ items)", {
  x: 0.7, y: 1.12, w: 4.0, h: 0.4,
  fontSize: 11, fontFace: "Arial", color: C.white, bold: true, margin: 0
});

const inventory = [
  { cat: "Stuffed Animals / Plush (20)", detail: "Varies in size (6\"–24\"), texture (fur, fleece, velour), and shape" },
  { cat: "Building Blocks (30)", detail: "LEGO sets, Mega Bloks, wooden blocks, magnetic tiles" },
  { cat: "Vehicles & Figurines (20)", detail: "Cars, trucks, action figures, dolls — rigid small objects" },
  { cat: "Balls & Sports (15)", detail: "Tennis, foam, rubber, soccer, basketball — varied grip" },
  { cat: "Books & Flat Items (15)", detail: "Board books, coloring books, puzzles in boxes" },
  { cat: "Art Supplies (15)", detail: "Crayons, markers, containers, play-doh tubs" },
  { cat: "Kitchen / Pretend Play (15)", detail: "Plastic food, utensils, pots, tool sets" },
  { cat: "Miscellaneous (20+)", detail: "Instruments, costumes, shoes, blankets, bins, baskets" },
];

inventory.forEach((item, i) => {
  const yPos = 1.6 + i * 0.45;
  slide7.addText(item.cat, { x: 0.7, y: yPos, w: 4.0, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  slide7.addText(item.detail, { x: 0.7, y: yPos + 0.2, w: 4.0, h: 0.2, fontSize: 9, fontFace: "Calibri", color: C.muted, margin: 0 });
});

// Right column: Scene variation strategy
slide7.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.4, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
slide7.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.4, h: 0.4, fill: { color: C.secondary } });
slide7.addText("SCENE VARIATION AXES", {
  x: 5.3, y: 1.12, w: 4.0, h: 0.4,
  fontSize: 11, fontFace: "Arial", color: C.white, bold: true, margin: 0
});

const variations = [
  { axis: "Object Count", desc: "Vary from 5 (sparse) to 40+ (cluttered) items per scene reset" },
  { axis: "Scatter Pattern", desc: "Concentrated cluster, spread across room, under/behind furniture" },
  { axis: "Surface Variation", desc: "Floor, couch, table, shelf, inside bins — different heights" },
  { axis: "Object Mix", desc: "Rotate object subsets daily; introduce 10 new items every 2 days" },
  { axis: "Tidying Destination", desc: "Vary target bins: toy chest, shelves, labeled baskets, bins" },
  { axis: "Lighting Conditions", desc: "Morning natural light vs. afternoon vs. overhead artificial" },
  { axis: "Obstacle Complexity", desc: "Clear path → furniture obstacles → nested/stacked objects" },
  { axis: "Reset Style", desc: "'Quick scatter' vs. 'staged mess' vs. 'realistic play aftermath'" },
];

variations.forEach((v, i) => {
  const yPos = 1.6 + i * 0.45;
  slide7.addText(v.axis, { x: 5.3, y: yPos, w: 4.0, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.secondary, bold: true, margin: 0 });
  slide7.addText(v.desc, { x: 5.3, y: yPos + 0.2, w: 4.0, h: 0.2, fontSize: 9, fontFace: "Calibri", color: C.muted, margin: 0 });
});

slide7.addNotes(`[~5 minutes — Inventory & Scene Variation Strategy]

On the left is the inventory plan. I've organized 150-plus objects across 8 categories, chosen to maximize diversity along several dimensions: size — from small crayons to large stuffed animals; material — soft plush, rigid plastic, deformable foam; shape — balls, blocks, flat books, irregular figurines; and weight — from lightweight toys to heavier play-doh tubs.

The category breakdown is designed so the robot encounters fundamentally different grasping challenges. Stuffed animals require different grip strategies than building blocks. Balls roll. Books are flat and slippery. Kitchen toys have handles and odd shapes. This diversity is what will make the trained policy generalizable.

For sourcing, I'd recommend a mix of Target and Amazon for new items — gives us consistent quality and easy re-ordering if something breaks — plus thrift stores for unusual or varied items at lower cost. Total budget estimate is in the range of $500-800, which I can detail in a separate inventory spreadsheet.

On the right are the 8 axes of scene variation. This is how we ensure the dataset isn't just "different toys in the same pile on the floor." Each axis represents a dimension we'll systematically vary:

Object count — from sparse 5-item scenes to cluttered 40-plus item messes. The policy needs to handle both.

Scatter pattern — sometimes toys are in a concentrated pile, sometimes spread across the entire room, sometimes hidden under furniture.

Surface variation — toys on the floor, couch, table, shelves. Different heights force different approach trajectories.

Object mix — we rotate subsets daily and introduce 10 new items every 2 days, so the robot never sees the same exact scene twice.

Tidying destination — we vary where toys go: a toy chest, labeled bins, shelves. This teaches flexible placement, not just "dump everything in one bin."

Lighting, obstacle complexity, and reset style round it out. The key principle is that we're not just collecting more data — we're collecting strategically diverse data.`);

// ============================================================
// SLIDE 8: 2-Week Schedule
// ============================================================
let slide8 = pres.addSlide();
slide8.background = { color: C.lightBg };

slide8.addText("2-Week Collection Schedule", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Week 1 & 2 table
const tableHeader = [
  { text: "Day", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, align: "center" } },
  { text: "Focus Theme", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, align: "center" } },
  { text: "Object Count", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, align: "center" } },
  { text: "Scene Resets", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, align: "center" } },
  { text: "Key Variation", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 9, align: "center" } },
];

const cellOpt = (bg) => ({ fill: { color: bg || C.white }, fontSize: 8.5, fontFace: "Calibri", color: C.text, align: "center" });
const altBg = "F1F5F9";

const rows = [
  tableHeader,
  ["1", "Setup & Baseline", "10–15 items", "4–5 / shift", "Simple floor scatter, soft toys only"].map((t, i) => ({ text: t, options: cellOpt() })),
  ["2", "Baseline (Mixed)", "15–20 items", "5–6 / shift", "Add blocks + vehicles, floor only"].map((t, i) => ({ text: t, options: cellOpt(altBg) })),
  ["3", "Surface Variation", "15–20 items", "5–6 / shift", "Couch, table surfaces, different heights"].map((t, i) => ({ text: t, options: cellOpt() })),
  ["4", "Increased Clutter", "25–30 items", "4–5 / shift", "Dense scatter, overlapping objects"].map((t, i) => ({ text: t, options: cellOpt(altBg) })),
  ["5", "Destination Variety", "20–25 items", "5–6 / shift", "Multiple bins, shelves, toy chest targets"].map((t, i) => ({ text: t, options: cellOpt() })),
  ["6–7", "WEEKEND — Data QA, review footage, adjust plan for Week 2", "", "", ""].map((t, i) => ({ text: t, options: { ...cellOpt("E8F5E9"), colspan: i === 1 ? 4 : undefined, align: i === 1 ? "left" : "center" } })),
  ["8", "New Object Rotation", "25–30 items", "6–7 / shift", "Swap 50% objects, add art supplies + books"].map((t, i) => ({ text: t, options: cellOpt() })),
  ["9", "Obstacle Navigation", "20–25 items", "5–6 / shift", "Furniture obstacles, under-table items"].map((t, i) => ({ text: t, options: cellOpt(altBg) })),
  ["10", "Max Complexity", "35–40+ items", "4–5 / shift", "Full inventory, nested + stacked items"].map((t, i) => ({ text: t, options: cellOpt() })),
  ["11", "Speed Runs", "15–20 items", "8–10 / shift", "Fast resets, aim for max episodes/day"].map((t, i) => ({ text: t, options: cellOpt(altBg) })),
  ["12", "Edge Cases + Wrap", "Varies", "6–8 / shift", "Unusual objects, difficult placements, final gaps"].map((t, i) => ({ text: t, options: cellOpt() })),
];

slide8.addTable(rows, {
  x: 0.5, y: 1.0, w: 9.0,
  colW: [0.6, 1.8, 1.3, 1.3, 4.0],
  border: { pt: 0.5, color: "DEE2E6" },
  rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
});

// Daily protocol box
slide8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.9, w: 9.0, h: 0.55, fill: { color: C.white }, shadow: makeShadow() });
slide8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.9, w: 0.07, h: 0.55, fill: { color: C.accent } });
slide8.addText([
  { text: "Daily Protocol: ", options: { bold: true, color: C.navy, fontSize: 10 } },
  { text: "30-min setup → 6.5 hrs collection → 30-min reset & log → 30-min EOD review. Each scene reset takes ~15 min and is logged (photo + object list) for reproducibility.", options: { color: C.muted, fontSize: 10 } }
], { x: 0.75, y: 4.9, w: 8.6, h: 0.55, fontFace: "Calibri", margin: 0, valign: "middle" });

slide8.addNotes(`[~5 minutes — Day-by-Day Schedule]

Here's the detailed day-by-day plan. The key design principle is progressive complexity — we start simple and ramp up.

Day 1 is setup and baseline. We start with just 10 to 15 soft toys on the floor. This is intentionally easy — it lets us calibrate the robot, verify the teleoperation link works reliably in the home, establish our reset workflow, and get baseline performance data. If something goes wrong with the setup, we'd rather discover it on a simple scene than a complex one.

Days 2 through 5 progressively add complexity: mixed object types, different surfaces, more clutter, and varied tidying destinations. By the end of week 1, we should have solid coverage of the core scenarios.

Days 6 and 7 are the weekend — and this is a critical checkpoint, not just time off. We'll review the footage and data from week 1, identify any gaps in coverage, check for data quality issues like we saw in the logistics dataset, and adjust the week 2 plan accordingly. For example, if we notice the robot struggles with balls rolling away, we might add more ball-heavy scenes in week 2.

Week 2 pushes further. Day 8 swaps out 50% of the objects — this is when the art supplies, books, and more unusual items come in. Day 9 adds furniture obstacles and under-table retrievals. Day 10 is maximum complexity — full inventory, nested and stacked objects, the hardest scenes we can create.

Day 11 is interesting — speed runs. We reduce object count back to 15-20 but aim for maximum resets per shift, targeting 8-10 scene resets. This is inspired by the logistics data where speed-mode episodes were our best performers. We want high-throughput, efficient tidying data.

Day 12 wraps up with edge cases and gap-filling — whatever the data review tells us we're missing.

The daily protocol at the bottom is important: 30 minutes of setup each morning, 6.5 hours of actual collection, then 30 minutes for reset and logging, and 30 minutes for end-of-day review. Each scene reset takes about 15 minutes and is documented with a photo and object list. This logging is critical for reproducibility — if we find later that certain scenes produced great training data, we need to be able to recreate them.`);

// ============================================================
// SLIDE 9: Logistics & Risk Mitigation
// ============================================================
let slide9 = pres.addSlide();
slide9.background = { color: C.lightBg };

slide9.addText("Logistics, QA & Risk Mitigation", {
  x: 0.6, y: 0.3, w: 8.8, h: 0.6,
  fontSize: 26, fontFace: "Georgia", color: C.navy, bold: true, margin: 0
});

// Three column cards
const logCards = [
  {
    title: "PRE-COLLECTION (Days 1–3)",
    color: C.accent,
    items: [
      "Source inventory: Target, Amazon, thrift stores for variety and cost efficiency",
      "Prepare scene reset checklists with photo templates for each variation",
      "Transport robot + equipment; verify WiFi, charging, and teleoperation link at home",
      "Create shared drive for daily logs, photos, and episode metadata",
      "Conduct dry run: 2-hr test session to calibrate cycle times and reset workflow",
    ]
  },
  {
    title: "DURING COLLECTION",
    color: C.secondary,
    items: [
      "Start each day with quick calibration check + photo of starting scene",
      "Rotate scene variation per schedule; log every reset with timestamp + photo",
      "Monitor data quality live — flag incomplete episodes, robot errors",
      "Mid-day check-in: review episode count vs. daily target, adjust reset frequency",
      "End-of-day debrief: upload data, note anomalies, prep next day's scene theme",
    ]
  },
  {
    title: "RISK MITIGATION",
    color: C.red,
    items: [
      "Robot downtime: Keep spare parts + remote support contact; buffer 1 day in schedule",
      "Data quality: Automated QA script to detect truncated episodes (<2s) daily",
      "Network issues: Local data backup drive; sync to cloud each evening",
      "Home access: Confirm schedule with homeowner; identify backup home if needed",
      "Object damage: Budget for replacements; keep receipts for inventory tracking",
    ]
  }
];

logCards.forEach((card, i) => {
  const xPos = 0.4 + i * 3.15;
  slide9.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.0, w: 3.0, h: 4.3, fill: { color: C.white }, shadow: makeShadow() });
  slide9.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.0, w: 3.0, h: 0.4, fill: { color: card.color } });
  slide9.addText(card.title, { x: xPos + 0.15, y: 1.02, w: 2.7, h: 0.4, fontSize: 9.5, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
  
  card.items.forEach((item, j) => {
    const yItem = 1.55 + j * 0.72;
    slide9.addText([
      { text: item, options: { bullet: true, fontSize: 9.5, color: C.text } }
    ], { x: xPos + 0.15, y: yItem, w: 2.7, h: 0.65, fontFace: "Calibri", margin: 0 });
  });
});

slide9.addNotes(`[~4 minutes — Logistics & Risk Mitigation, then Q&A]

This last slide covers the practical logistics — the things that can make or break a data collection even if the plan is good on paper.

Pre-collection — the next three days before our manager meeting: I need to finalize the inventory order, prepare the scene reset checklists with photo templates, and critically, do a dry run at the home. That dry run is non-negotiable. We need to verify the WiFi is stable enough for teleoperation, that we have adequate charging, and that the robot can physically navigate the space. The logistics dataset showed us that early episodes on any new setup are slower — I want to burn through that ramp-up before the real collection starts.

During collection, the key discipline is logging. Every scene reset gets a timestamp, a photo, and an object list. This serves two purposes: it makes the data reproducible, and it lets us do mid-collection QA. The mid-day check-in is important — if we're falling behind on episode count, we can adjust reset frequency. If we notice the robot struggling with a particular scenario, we can note it for the end-of-day debrief.

On risk mitigation — I've identified five main risks:

Robot downtime is the biggest. We're an hour from the office with one robot. If it goes down, we lose a full day plus travel time. Mitigation: keep common spare parts on site and have remote support contact on speed dial. I've also built one buffer day into the schedule.

Data quality — we learned from the logistics data that truncated episodes can pollute the training set. I want to run an automated QA script daily that flags episodes under 2 seconds.

Network issues — if the home WiFi drops during teleoperation, we could lose episode data. We'll keep a local backup drive and sync to cloud each evening.

Home access — we need to confirm the full 2-week schedule with the homeowner upfront, and I'd recommend identifying a backup home just in case.

And object damage — toys will break. We need a replacement budget and receipts for inventory tracking.

That's the full plan. I'd welcome your feedback, especially on the scene variation strategy and whether the complexity ramp feels right. I'm also happy to detail the inventory budget or discuss how we'd adjust if the dry run reveals constraints I haven't anticipated.

[Open for Q&A — remaining time]`);

// ============================================================
// Write file
// ============================================================
pres.writeFile({ fileName: "/home/claude/figure_case_study.pptx" })
  .then(() => console.log("Presentation saved!"))
  .catch(err => console.error(err));
