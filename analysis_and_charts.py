"""
Figure AI Case Study — Data Analysis & Chart Generation
=========================================================

This script analyzes 7,835 teleoperation episodes from a conveyor belt package
placement task and generates 5 matplotlib charts used in the final presentation.

KEY ASSUMPTION (applies to ALL numbers in this analysis):
    Each episode involves ~10 packages. The raw data only gives total episode
    length in seconds, NOT individual package times. We divide by 10 to get an
    "implied cycle time per package." This assumption is supported by:
      - The case study prompt sets a target of 3.5 sec/package
      - Mean episode length is ~37s, which / 10 = 3.7s (close to target)
      - The ~195 episodes that DO specify "10_packages" in their name have
        similar mean lengths to those that don't
    If the actual package count differs, all cycle time numbers shift proportionally.

DATA PIPELINE:
    1. Load raw Excel -> 7,835 rows (3 columns: dataset name, episode ID, length)
    2. Parse dataset names with regex to extract: date, config, pilot, flags
    3. Exclude 195 early pilot episodes with explicit package/box counts
       (different task structure) -> 7,640 "main" episodes
    4. Compute implied_cycle = episode_length / 10
    5. Group by various dimensions and plot
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend (no display needed, just saving PNGs)
import matplotlib.pyplot as plt
import numpy as np
import re

# =============================================================================
# STEP 1: LOAD AND CLEAN DATA
# =============================================================================

# The Excel file has a blank first row, so skiprows=1 to get the actual header.
# Column 'Unnamed: 0' is an empty index column from the export — we drop it.
df = pd.read_excel('/mnt/user-data/uploads/_Figure_Case_Study__Logistics_collect_data.xlsx', skiprows=1)
df = df.drop(columns=['Unnamed: 0'])
df.columns = ['Dataset_name', 'Episode_ID', 'Episode_length_sec']

# Some rows have non-numeric episode lengths (likely formatting artifacts) — coerce to NaN
df['Episode_length_sec'] = pd.to_numeric(df['Episode_length_sec'], errors='coerce')
# Drop rows with no dataset name (empty rows at the end of the file)
df = df.dropna(subset=['Dataset_name'])

# =============================================================================
# STEP 2: PARSE DATASET NAMES
# =============================================================================
# Dataset names encode metadata in underscore-separated fields, e.g.:
#   "20251114_abc_abc_f03_aaaaa_c-100_Pilot4_abc"
#   "20251120_abc_abc_f03_aaaaa_c-111_Pilot1_hq_no_bbera"
#
# The format is NOT perfectly consistent across dates, so we use regex rather
# than positional parsing. Fields extracted:
#   - date: first 8 characters (YYYYMMDD)
#   - pilot: regex for "Pilot" followed by digits (Pilot1-Pilot5)
#   - config: regex for "c-" followed by digits (c-100 through c-198)
#             we take the LAST match because some names have c-XXX in the
#             episode ID portion too
#   - quality: presence of "_hq" in the name -> "hq", otherwise "abc"
#   - no_bb: presence of "no_bb" substring (the "bbera" flag — meaning unknown)
#   - speed: presence of "speed" substring (speed teleoperation mode)

def parse_dataset_v2(name):
    info = {}
    parts = name.split('_')
    info['date'] = parts[0]

    # Find pilot ID — always formatted as "PilotN"
    pilot_match = re.search(r'(Pilot\d+)', name)
    info['pilot'] = pilot_match.group(1) if pilot_match else None

    # Find robot config — always formatted as "c-NNN"
    # Take last match to avoid false positives from episode ID substring
    config_matches = re.findall(r'(c-\d+)', name)
    info['config'] = config_matches[-1] if config_matches else None

    # Quality flag: "hq" appears explicitly in some dataset names
    info['quality'] = 'hq' if '_hq' in name else 'abc'

    # Special operational flags
    info['no_bb'] = 'no_bb' in name.lower()     # "no_bbera" — meaning unclear
    info['speed'] = 'speed' in name.lower()       # speed teleoperation mode

    return info

parsed = df['Dataset_name'].apply(parse_dataset_v2).apply(pd.Series)
df = pd.concat([df, parsed], axis=1)

# =============================================================================
# STEP 3: FILTER TO MAIN DATASET
# =============================================================================
# 195 early episodes (Oct 20) have explicit "1_package_1_box" or "10_packages"
# in their names. These are from an initial pilot with a different task structure.
# We exclude them so the main analysis covers a consistent task format.
#
# Result: 7,640 episodes across 10 dates, 15 configs, 5 pilots.
main = df[~df['Dataset_name'].str.contains('package|box', case=False)].copy()

# =============================================================================
# STEP 4: COMPUTE IMPLIED CYCLE TIME
# =============================================================================
# This is THE core metric. We divide total episode length by 10 (assumed packages).
# Target from the case study prompt: 3.5 seconds per package.
# So target episode length = 3.5 x 10 = 35 seconds.
main['implied_cycle'] = main['Episode_length_sec'] / 10

# =============================================================================
# CHART STYLING
# =============================================================================
# Consistent color palette across all 5 charts — "Midnight Executive" theme
# chosen to match the PowerPoint deck.
plt.rcParams['font.family'] = 'sans-serif'
colors = {
    'primary': '#1E2761',    # Navy — main data lines and titles
    'secondary': '#065A82',  # Teal-blue — bars that are ABOVE target
    'accent': '#0D9488',     # Teal-green — bars that are BELOW target (good)
    'light': '#CADCFC',      # Ice blue — background accents
    'bg': '#F8FAFB',         # Off-white — chart background
    'text': '#1E293B',       # Dark slate — axis labels
    'target': '#E74C3C'      # Red — the 3.5s target line
}

# =============================================================================
# CHART 1: CYCLE TIME OVER THE COLLECTION PERIOD (Timeline)
# =============================================================================
# PURPOSE: Show whether performance improved over the 10-day collection window.
#
# HOW IT'S COMPUTED:
#   - Group all 7,640 episodes by their date (10 unique dates)
#   - For each date, compute mean and median of implied_cycle (= episode_length / 10)
#   - Plot both lines so the reader can see if the mean is pulled by outliers
#
# WHY BOTH MEAN AND MEDIAN:
#   The distribution is right-skewed (long tail of slow episodes). The median
#   is more robust to outliers, so if mean >> median on a given day, it means
#   a few very slow episodes are inflating the average.
#
# THE RED SHADED AREA shows days where the mean is above the 3.5s target.
#
# KEY FINDING: Nov 22 (2.98s) and Nov 26 (3.06s) both beat the target,
# showing a clear improvement trend over the collection period.

fig, ax = plt.subplots(figsize=(10, 5))
date_stats = main.groupby('date')['implied_cycle'].agg(['mean', 'median', 'count'])
date_stats.index = pd.to_datetime(date_stats.index, format='%Y%m%d')

ax.plot(date_stats.index, date_stats['mean'], 'o-',
        color=colors['primary'], linewidth=2, markersize=8, label='Mean Cycle Time')
ax.plot(date_stats.index, date_stats['median'], 's--',
        color=colors['accent'], linewidth=2, markersize=7, label='Median Cycle Time')
ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8,
           label='Target (3.5s)')

# Shade the gap between mean and target where mean > target
ax.fill_between(date_stats.index, 3.5, date_stats['mean'],
                where=date_stats['mean'] > 3.5, alpha=0.1, color=colors['target'])

ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_xlabel('')
ax.set_title('Average Cycle Time Over Collection Period',
             fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, alpha=0.3)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('/home/shriya/chart1_timeline.png', dpi=200, bbox_inches='tight')
plt.close()

# =============================================================================
# CHART 2: TELEOPERATOR PERFORMANCE COMPARISON (Bar chart)
# =============================================================================
# PURPOSE: Compare the 5 operators (Pilot1-Pilot5) to identify who's fastest
# and quantify the skill gap.
#
# HOW IT'S COMPUTED:
#   - Group by pilot -> compute mean, median, std, count of implied_cycle
#   - Sort by mean (best to worst)
#   - Error bars = standard error of mean (std / sqrt(n)), NOT raw std dev.
#     This shows confidence in the mean estimate, not episode-level variation.
#
# BAR COLORING LOGIC:
#   - Teal-green (accent) if mean < 3.5s (under target) — none currently qualify
#   - Blue (secondary) if mean >= 3.5s (above target)
#
# WHY THIS MATTERS:
#   A 9% gap between best (Pilot2/3 at ~3.55s) and worst (Pilot5 at 3.86s)
#   on the SAME hardware suggests coaching/technique sharing could close the gap.
#   The error bars are small relative to the differences, so these gaps are
#   statistically meaningful (not just noise).

fig, ax = plt.subplots(figsize=(9, 5))
pilot_stats = main.groupby('pilot')['implied_cycle'].agg(['mean', 'median', 'std', 'count'])
pilot_stats = pilot_stats.sort_values('mean')  # Best operator first

bars = ax.bar(pilot_stats.index, pilot_stats['mean'],
              color=[colors['accent'] if m < 3.5 else colors['secondary']
                     for m in pilot_stats['mean']],
              edgecolor='white', linewidth=1.5, width=0.6)

# Error bars: standard error = std / sqrt(n)
# This narrows with more data — Pilot3 (n=2753) has a tighter estimate than Pilot4 (n=848)
ax.errorbar(pilot_stats.index, pilot_stats['mean'],
            yerr=pilot_stats['std'] / np.sqrt(pilot_stats['count']),
            fmt='none', color=colors['text'], capsize=5, alpha=0.7)

ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8,
           label='Target')

# Annotate each bar with its value and sample size
for i, (idx, row) in enumerate(pilot_stats.iterrows()):
    ax.text(i, row['mean'] + 0.08, f'{row["mean"]:.2f}s',
            ha='center', fontsize=11, fontweight='bold', color=colors['text'])
    ax.text(i, row['mean'] - 0.15, f'n={int(row["count"])}',
            ha='center', fontsize=9, color='white')

ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Teleoperator Performance Comparison',
             fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/shriya/chart2_operators.png', dpi=200, bbox_inches='tight')
plt.close()

# =============================================================================
# CHART 3: PERFORMANCE BY ROBOT CONFIGURATION (Horizontal bar chart)
# =============================================================================
# PURPOSE: Compare the 15+ robot configs to identify which hardware/software
# settings produce the fastest cycle times.
#
# HOW IT'S COMPUTED:
#   - Group by config -> compute mean implied_cycle and episode count
#   - FILTER: only include configs with >= 50 episodes. This excludes c-128 (13 eps),
#     c-129 (11 eps). Small samples have unreliable means — e.g., c-129's "3.53s"
#     is based on just 11 episodes with a std of 15.3s, so its true mean could
#     be anywhere from 1s to 6s.
#   - Sort by mean (best at top)
#
# BAR COLORING: same logic as Chart 2 — green if under 3.5s target.
#
# KEY FINDING:
#   c-198 (3.21s, n=565) and c-144 (3.41s, n=2086) are the only configs that
#   beat the target with substantial sample sizes. c-198 is the newest config
#   (only active Nov 25-26), suggesting recent improvements are working.
#   c-103 (4.01s) and c-106 (4.11s) are the worst — both older configs.

fig, ax = plt.subplots(figsize=(10, 5))
config_stats = main.groupby('config')['implied_cycle'].agg(['mean', 'count'])
config_stats = config_stats[config_stats['count'] >= 50].sort_values('mean')

bar_colors = [colors['accent'] if m < 3.5 else colors['secondary']
              for m in config_stats['mean']]
bars = ax.barh(config_stats.index, config_stats['mean'],
               color=bar_colors, edgecolor='white', height=0.6)

ax.axvline(x=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8,
           label='Target')

# Annotate with value and sample size so the reader can gauge reliability
for i, (idx, row) in enumerate(config_stats.iterrows()):
    ax.text(row['mean'] + 0.05, i,
            f'{row["mean"]:.2f}s (n={int(row["count"])})',
            va='center', fontsize=9, color=colors['text'])

ax.set_xlabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Performance by Robot Configuration',
             fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/shriya/chart3_configs.png', dpi=200, bbox_inches='tight')
plt.close()

# =============================================================================
# CHART 4: DISTRIBUTION OF CYCLE TIMES (Histogram)
# =============================================================================
# PURPOSE: Show the full shape of the data — are most episodes clustered near
# the target, or is there a long tail of slow episodes inflating the mean?
#
# HOW IT'S COMPUTED:
#   - Take all 7,640 implied_cycle values (episode_length / 10)
#   - Plot a histogram with 50 bins — enough granularity to see the shape
#     without being too noisy
#   - Overlay two vertical lines:
#       Red dashed = 3.5s target
#       Navy solid = actual mean (3.70s)
#   - Annotate with the % of episodes that fall LEFT of the target line
#
# WHY 50 BINS:
#   The data ranges from ~0s to ~11.6s. With 50 bins, each bin is ~0.23s wide.
#   This is granular enough to see the peak (~3.5s) and the right tail clearly.
#   Fewer bins (20) would hide the peak shape; more (100) would be noisy.
#
# KEY FINDING:
#   The distribution peaks around 3.3-3.7s (right at the target zone), but has
#   a clear right tail extending past 6s. 41% of episodes are already under
#   target — the challenge is reducing the slow tail, not shifting the whole
#   distribution.
#
# OUTLIER NOTE: 44 episodes are < 5s total (< 0.5s/pkg implied). These are
# almost certainly aborted or erroneously recorded — you can't place 10 packages
# in under 5 seconds. They appear as the tiny blip at the far left.

fig, ax = plt.subplots(figsize=(9, 5))
ax.hist(main['implied_cycle'], bins=50,
        color=colors['secondary'], edgecolor='white', alpha=0.85)

ax.axvline(x=3.5, color=colors['target'], linestyle='--', linewidth=2.5,
           label='Target (3.5s)')
ax.axvline(x=main['implied_cycle'].mean(), color=colors['primary'],
           linestyle='-', linewidth=2,
           label=f'Mean ({main["implied_cycle"].mean():.2f}s)')

# Calculate and annotate the % under target
pct_under = (main['implied_cycle'] < 3.5).mean() * 100
ax.text(3.5, ax.get_ylim()[1]*0.9, f'  {pct_under:.0f}% under target',
        fontsize=11, color=colors['target'], fontweight='bold')

ax.set_xlabel('Implied Cycle Time per Package (sec)', fontsize=12, color=colors['text'])
ax.set_ylabel('Frequency', fontsize=12, color=colors['text'])
ax.set_title('Distribution of Cycle Times',
             fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/shriya/chart4_distribution.png', dpi=200, bbox_inches='tight')
plt.close()

# =============================================================================
# CHART 5: SPEED MODE IMPACT (Bar chart — 2 bars)
# =============================================================================
# PURPOSE: Quantify how much faster "speed mode" teleoperation is vs. normal.
#
# HOW IT'S COMPUTED:
#   - Split episodes into two groups based on whether "speed" appears in the
#     dataset name (157 speed episodes vs. 7,483 normal)
#   - Compute mean implied_cycle for each group
#   - Plot side-by-side bars
#
# WHY THIS MATTERS:
#   Speed mode averages 2.87s/pkg vs. 3.72s/pkg normal — a 23% improvement.
#   This proves the 3.5s target is physically achievable with the current hardware.
#   The question becomes: can we make "speed mode" the default, or is there a
#   quality/accuracy tradeoff we're not seeing in this data?
#
# CAVEAT: Only 157 speed episodes (2% of data), all on a single config (c-198)
# by a subset of pilots. The speed advantage could partially be a config effect
# (c-198 is already the fastest config at 3.21s even in normal mode). A proper
# A/B test — same config, same pilot, speed vs. normal — would isolate the true
# speed mode effect.

fig, ax = plt.subplots(figsize=(7, 4.5))
speed_data = main.groupby('speed')['implied_cycle'].agg(['mean', 'std', 'count'])
speed_data.index = ['Normal', 'Speed Mode']

bars = ax.bar(speed_data.index, speed_data['mean'],
              color=[colors['secondary'], colors['accent']],
              edgecolor='white', width=0.5)

# Annotate bars with their values
for i, (idx, row) in enumerate(speed_data.iterrows()):
    ax.text(i, row['mean'] + 0.05, f'{row["mean"]:.2f}s',
            ha='center', fontsize=12, fontweight='bold', color=colors['text'])

ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8,
           label='Target')

ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Impact of Speed Mode',
             fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/shriya/chart5_speed.png', dpi=200, bbox_inches='tight')
plt.close()

print("All charts saved successfully!")
