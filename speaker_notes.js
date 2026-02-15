const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, LevelFormat } = require('docx');
const fs = require('fs');

const notes = [
  {
    slide: "Slide 1 — Title",
    time: "~2 minutes",
    content: `Good morning/afternoon everyone. Thank you for taking the time to review this case study with me.

Today I'll walk you through two parts. First, a data analysis of our logistics teleoperation dataset — over 7,800 episodes of a robot placing packages onto a conveyor belt. I'll share key patterns, performance drivers, and recommendations for improving the autonomous policy we're training from this data.

Second, I'll present an operational plan for a 2-week data collection effort in a residential home, where we'll be teaching a robot to tidy up a playroom.

Let's start with the data analysis.`
  },
  {
    slide: "Slide 2 — Dataset Overview & Key Metrics",
    time: "~4 minutes",
    content: `Let me orient you on the dataset. We collected 7,835 teleoperation episodes across 10 collection days in November 2025. After filtering out 195 early pilot episodes with different task setups, our main analysis covers 7,640 episodes spanning 15 robot configurations — labeled c-100 through c-198 — and 5 teleoperators, Pilot 1 through 5.

Each episode represents a full run of placing packages onto the conveyor belt with barcodes facing down. The episode length captures the total time, and since most episodes involve 10 packages, I've derived an "implied cycle time" by dividing episode length by 10.

Looking at the headline numbers: our current average cycle time is 3.70 seconds per package, which is about 6% above the 3.5-second target. However — and this is important — 41% of episodes already meet the target. And when operators use "speed mode," the average drops to 2.87 seconds, well below target. So the 3.5-second goal is clearly achievable — the question is what's dragging the average up and how we filter or improve the training data accordingly.

Let me dig into the trends.`
  },
  {
    slide: "Slide 3 — Cycle Time Trends & Distribution",
    time: "~5 minutes",
    content: `On the left, you can see cycle time plotted over the collection period. The trend is clearly downward — we started around 5.5 seconds in mid-October, and by the last two collection days — November 22nd and 26th — we're averaging under 3 seconds. That's a dramatic improvement.

A few things are driving this. First, there's a learning curve — operators get faster with practice. Second, we introduced newer robot configurations over time, and those newer configs tend to perform better. I'll break that down on the next slide.

On the right is the distribution of all cycle times. You can see the bulk of episodes cluster between 2.5 and 5 seconds, with a peak around 3.5 to 4 seconds. The red dashed line is our target. 41% of episodes — everything to the left of that line — already meet it.

There are two tails worth noting. On the left, we have 44 episodes under 0.5 seconds — these are almost certainly aborted or erroneously recorded episodes, not real completions. On the right, we have a long tail extending past 8 seconds, which likely represents episodes with errors, pauses, or difficult package orientations.

Both tails represent data quality issues. My recommendation is to filter these out before training — they'd teach the policy either incomplete trajectories or inefficient ones. I'd suggest a window of roughly 2 to 6 seconds per package as a clean training range, though that threshold should be validated with the robotics team.`
  },
  {
    slide: "Slide 4 — Performance by Operator & Configuration",
    time: "~6 minutes",
    content: `This is where it gets interesting. On the left, we have operator performance. Pilot 3 and Pilot 2 are our fastest operators at 3.56 and 3.60 seconds respectively — both very close to target. Pilot 5 and Pilot 1 are the slowest at 3.86 and 3.90 seconds. That's roughly a 9% gap between best and worst.

Now, 9% might not sound like much, but when you're trying to close a 6% gap to target, operator technique matters. I'd recommend having Pilot 2 and 3 share their approach with the slower operators — there may be specific teleoperation techniques or strategies they use that others could adopt.

On the right is the more impactful finding: robot configuration variance. This is a much bigger driver than operator skill. Config c-198, our newest, averages 3.21 seconds — well under target. Config c-144, our largest dataset with over 2,000 episodes, hits 3.41 seconds. But older configs like c-103 and c-106 are stuck above 4 seconds.

There's an important caveat here: configs were introduced sequentially, not tested in parallel. So the improvement from c-103 to c-198 could partly reflect operator learning rather than purely hardware or software changes. We'd need controlled testing — same operator, same day, different configs — to fully separate these effects.

I also looked at learning curves within individual configs. Config c-103 improved by 11.7 seconds comparing its first 50 episodes to its last 50. Config c-144 improved by 6.5 seconds. This suggests early episodes in any config are significantly slower as operators ramp up. For training data, we may want to exclude the first N episodes per config to avoid teaching the policy suboptimal early behavior.

One more data point from the cross-tabulation: Pilot 2 on config c-144 averages 3.06 seconds, while Pilot 5 on the same config averages 4.00 seconds. Same hardware, nearly 1 second gap — that confirms operator skill and config interact in ways we should understand better.`
  },
  {
    slide: "Slide 5 — Insights & Next Steps",
    time: "~6 minutes",
    content: `Let me highlight three specific findings before moving to recommendations.

First, speed mode. About 157 episodes were collected in an explicit "speed mode," and they average 2.87 seconds — 23% faster than normal operation. This is our strongest evidence that the 3.5-second target is not only achievable but beatable. The question for the team is whether speed-mode data should be weighted more heavily in training, or whether it introduces behaviors that sacrifice reliability — like missed barcode orientations.

Second, the "no_bbera" flag. About 1,600 episodes carry this flag in their dataset name, but I don't know what "bbera" refers to — I initially assumed bounding boxes, but that turned out to be incorrect. What I can say is that the flag correlates with measurable performance differences: on some configs it's associated with faster episodes, on others slower. Before we draw any conclusions, we need to clarify with the team what this variable actually controls. Once we know that, a controlled A/B test would tell us whether to standardize it.

Third, outliers. We have 44 episodes under 5 seconds and 6 over 80 seconds. The short ones are almost certainly incomplete — you physically cannot place 10 packages in under 5 seconds. The long ones likely involve robot errors, operator breaks, or connectivity issues. Both should be flagged and excluded from training data.

Now for recommended additional analyses — and I want to emphasize these are things I think would meaningfully improve policy quality, not just academic exercises:

The most valuable would be a per-package breakdown. Right now, we only have total episode time. If we can decompose each episode into individual pick-place cycles, we could identify which package positions or orientations are systematically slower. That would tell us where the policy needs the most training signal.

Second, correlating episode length with scan success rate. We're optimizing for speed, but the constraint is barcode-down placement. Understanding the quality-speed frontier would help us set the right training target — maybe 3.5 seconds is too aggressive if it leads to scan failures, or maybe we can push even faster.

Third, time-of-day effects. I noticed episodes around noon are about 3 seconds faster than episodes at 8 PM. This could reflect operator fatigue, lighting changes, or shift patterns. If late-shift data is systematically worse, we should either downweight it or adjust scheduling.

Fourth, controlled pairings. Run our best operators on our best configs — Pilot 2 and 3 on c-198 and c-144 — to establish a true performance ceiling. That tells us the upper bound of what's possible with current hardware.

And fifth, a data filtering strategy. Combining everything — exclude ramp-up episodes, outliers, non-HQ quality data — to create a curated training set. The current dataset has a lot of noise, and a cleaner subset might train a better policy even if it's smaller.`
  },
  {
    slide: "Slide 6 — Operational Plan Introduction",
    time: "~3 minutes",
    content: `Now let's shift to the second part: the operational plan for our 2-week playroom data collection.

The setup: we have one robot, one residential home about an hour from the office, and 8-hour shifts each day. The task is tidying up a playroom — picking up toys and putting them away. The critical requirement is dataset diversity — we need a wide variety of objects, placements, and tidying approaches to train a robust policy.

Let me walk you through how I'm thinking about this across three areas: what we need to buy, how we vary the scenes, and the day-by-day schedule.

A few key numbers to keep in mind: 10 working days across the 2 weeks, roughly 80 hours of total robot runtime, and my target is to collect data on at least 150 unique objects. The weekend in between gives us a natural checkpoint to review data quality and adjust the plan for week 2 — something I learned from the logistics dataset, where we saw performance change significantly over time.`
  },
  {
    slide: "Slide 7 — Inventory Acquisition & Scene Variation",
    time: "~5 minutes",
    content: `On the left is the inventory plan. I've organized 150-plus objects across 8 categories, chosen to maximize diversity along several dimensions: size — from small crayons to large stuffed animals; material — soft plush, rigid plastic, deformable foam; shape — balls, blocks, flat books, irregular figurines; and weight — from lightweight toys to heavier play-doh tubs.

The category breakdown is designed so the robot encounters fundamentally different grasping challenges. Stuffed animals require different grip strategies than building blocks. Balls roll. Books are flat and slippery. Kitchen toys have handles and odd shapes. This diversity is what will make the trained policy generalizable.

For sourcing, I'd recommend a mix of Target and Amazon for new items — gives us consistent quality and easy re-ordering if something breaks — plus thrift stores for unusual or varied items at lower cost. Total budget estimate is in the range of $500-800, which I can detail in a separate inventory spreadsheet.

On the right are the 8 axes of scene variation. This is how we ensure the dataset isn't just "different toys in the same pile on the floor." Each axis represents a dimension we'll systematically vary:

Object count — from sparse 5-item scenes to cluttered 40-plus item messes. The policy needs to handle both.

Scatter pattern — sometimes toys are in a concentrated pile, sometimes spread across the entire room, sometimes hidden under furniture.

Surface variation — toys on the floor, couch, table, shelves. Different heights force different approach trajectories.

Object mix — we rotate subsets daily and introduce 10 new items every 2 days, so the robot never sees the same exact scene twice.

Tidying destination — we vary where toys go: a toy chest, labeled bins, shelves. This teaches flexible placement, not just "dump everything in one bin."

Lighting, obstacle complexity, and reset style round it out. The key principle is that we're not just collecting more data — we're collecting strategically diverse data.`
  },
  {
    slide: "Slide 8 — 2-Week Collection Schedule",
    time: "~5 minutes",
    content: `Here's the detailed day-by-day plan. The key design principle is progressive complexity — we start simple and ramp up.

Day 1 is setup and baseline. We start with just 10 to 15 soft toys on the floor. This is intentionally easy — it lets us calibrate the robot, verify the teleoperation link works reliably in the home, establish our reset workflow, and get baseline performance data. If something goes wrong with the setup, we'd rather discover it on a simple scene than a complex one.

Days 2 through 5 progressively add complexity: mixed object types, different surfaces, more clutter, and varied tidying destinations. By the end of week 1, we should have solid coverage of the core scenarios.

Days 6 and 7 are the weekend — and this is a critical checkpoint, not just time off. We'll review the footage and data from week 1, identify any gaps in coverage, check for data quality issues like we saw in the logistics dataset, and adjust the week 2 plan accordingly. For example, if we notice the robot struggles with balls rolling away, we might add more ball-heavy scenes in week 2.

Week 2 pushes further. Day 8 swaps out 50% of the objects — this is when the art supplies, books, and more unusual items come in. Day 9 adds furniture obstacles and under-table retrievals. Day 10 is maximum complexity — full inventory, nested and stacked objects, the hardest scenes we can create.

Day 11 is interesting — speed runs. We reduce object count back to 15-20 but aim for maximum resets per shift, targeting 8-10 scene resets. This is inspired by the logistics data where speed-mode episodes were our best performers. We want high-throughput, efficient tidying data.

Day 12 wraps up with edge cases and gap-filling — whatever the data review tells us we're missing.

The daily protocol at the bottom is important: 30 minutes of setup each morning, 6.5 hours of actual collection, then 30 minutes for reset and logging, and 30 minutes for end-of-day review. Each scene reset takes about 15 minutes and is documented with a photo and object list. This logging is critical for reproducibility — if we find later that certain scenes produced great training data, we need to be able to recreate them.`
  },
  {
    slide: "Slide 9 — Logistics, QA & Risk Mitigation",
    time: "~4 minutes + Q&A",
    content: `This last slide covers the practical logistics — the things that can make or break a data collection even if the plan is good on paper.

Pre-collection — the next three days before our manager meeting: I need to finalize the inventory order, prepare the scene reset checklists with photo templates, and critically, do a dry run at the home. That dry run is non-negotiable. We need to verify the WiFi is stable enough for teleoperation, that we have adequate charging, and that the robot can physically navigate the space. The logistics dataset showed us that early episodes on any new setup are slower — I want to burn through that ramp-up before the real collection starts.

During collection, the key discipline is logging. Every scene reset gets a timestamp, a photo, and an object list. This serves two purposes: it makes the data reproducible, and it lets us do mid-collection QA. The mid-day check-in is important — if we're falling behind on episode count, we can adjust reset frequency. If we notice the robot struggling with a particular scenario, we can note it for the end-of-day debrief.

On risk mitigation — I've identified five main risks:

Robot downtime is the biggest. We're an hour from the office with one robot. If it goes down, we lose a full day plus travel time. Mitigation: keep common spare parts on site and have remote support contact on speed dial. I've also built one buffer day into the schedule.

Data quality — we learned from the logistics data that truncated episodes can pollute the training set. I want to run an automated QA script daily that flags episodes under 2 seconds.

Network issues — if the home WiFi drops during teleoperation, we could lose episode data. We'll keep a local backup drive and sync to cloud each evening.

Home access — we need to confirm the full 2-week schedule with the homeowner upfront, and I'd recommend identifying a backup home just in case.

And object damage — toys will break. We need a replacement budget and receipts for inventory tracking.

That's the full plan. I'd welcome your feedback, especially on the scene variation strategy and whether the complexity ramp feels right. I'm also happy to detail the inventory budget or discuss how we'd adjust if the dry run reveals constraints I haven't anticipated.

[Open for Q&A — remaining time]`
  }
];

// Build document
const children = [];

// Title
children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after: 100 },
  children: [new TextRun({ text: "FIGURE CASE STUDY", font: "Arial", size: 28, bold: true, color: "0D9488" })]
}));
children.push(new Paragraph({
  spacing: { after: 300 },
  children: [new TextRun({ text: "Speaker Notes — 35-Minute Presentation", font: "Arial", size: 48, bold: true, color: "1E2761" })]
}));

// Time overview
children.push(new Paragraph({
  spacing: { after: 200 },
  children: [new TextRun({ text: "Time Allocation Overview", font: "Arial", size: 28, bold: true, color: "1E2761" })]
}));

const timeRows = [
  ["Slide 1 — Title", "~2 min"],
  ["Slide 2 — Dataset Overview", "~4 min"],
  ["Slide 3 — Trends & Distribution", "~5 min"],
  ["Slide 4 — Operator & Config", "~6 min"],
  ["Slide 5 — Insights & Next Steps", "~6 min"],
  ["Slide 6 — Ops Plan Intro", "~3 min"],
  ["Slide 7 — Inventory & Variation", "~5 min"],
  ["Slide 8 — 2-Week Schedule", "~5 min"],
  ["Slide 9 — Logistics & Risk + Q&A", "~4 min"],
];

for (const [slide, time] of timeRows) {
  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: slide, font: "Arial", size: 22 }),
      new TextRun({ text: "  " + time, font: "Arial", size: 22, bold: true, color: "0D9488" }),
    ]
  }));
}

children.push(new Paragraph({ spacing: { after: 200 }, children: [] }));

// Each slide's notes
for (const note of notes) {
  // Divider line
  children.push(new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DEE2E6" } },
    spacing: { after: 200 },
    children: []
  }));

  // Slide heading
  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: note.slide, font: "Arial", size: 28, bold: true, color: "1E2761" })]
  }));

  // Time tag
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: note.time, font: "Arial", size: 22, bold: true, color: "0D9488" })]
  }));

  // Content paragraphs
  const paragraphs = note.content.split('\n\n');
  for (const para of paragraphs) {
    if (para.trim()) {
      children.push(new Paragraph({
        spacing: { after: 160, line: 300 },
        children: [new TextRun({ text: para.trim(), font: "Arial", size: 22, color: "1E293B" })]
      }));
    }
  }
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/speaker_notes.docx", buffer);
  console.log("Speaker notes doc saved!");
});
