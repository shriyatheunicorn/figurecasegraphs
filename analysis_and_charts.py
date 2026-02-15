import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import re

# Setup
df = pd.read_excel('/mnt/user-data/uploads/_Figure_Case_Study__Logistics_collect_data.xlsx', skiprows=1)
df = df.drop(columns=['Unnamed: 0'])
df.columns = ['Dataset_name', 'Episode_ID', 'Episode_length_sec']
df['Episode_length_sec'] = pd.to_numeric(df['Episode_length_sec'], errors='coerce')
df = df.dropna(subset=['Dataset_name'])

def parse_dataset_v2(name):
    info = {}
    parts = name.split('_')
    info['date'] = parts[0]
    pilot_match = re.search(r'(Pilot\d+)', name)
    info['pilot'] = pilot_match.group(1) if pilot_match else None
    config_matches = re.findall(r'(c-\d+)', name)
    info['config'] = config_matches[-1] if config_matches else None
    info['quality'] = 'hq' if '_hq' in name else 'abc'
    info['no_bb'] = 'no_bb' in name.lower()
    info['speed'] = 'speed' in name.lower()
    return info

parsed = df['Dataset_name'].apply(parse_dataset_v2).apply(pd.Series)
df = pd.concat([df, parsed], axis=1)
main = df[~df['Dataset_name'].str.contains('package|box', case=False)].copy()
main['implied_cycle'] = main['Episode_length_sec'] / 10

# Style
plt.rcParams['font.family'] = 'sans-serif'
colors = {'primary': '#1E2761', 'secondary': '#065A82', 'accent': '#0D9488', 
          'light': '#CADCFC', 'bg': '#F8FAFB', 'text': '#1E293B', 'target': '#E74C3C'}

# Chart 1: Performance over time
fig, ax = plt.subplots(figsize=(10, 5))
date_stats = main.groupby('date')['implied_cycle'].agg(['mean', 'median', 'count'])
date_stats.index = pd.to_datetime(date_stats.index, format='%Y%m%d')
ax.plot(date_stats.index, date_stats['mean'], 'o-', color=colors['primary'], linewidth=2, markersize=8, label='Mean Cycle Time')
ax.plot(date_stats.index, date_stats['median'], 's--', color=colors['accent'], linewidth=2, markersize=7, label='Median Cycle Time')
ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8, label='Target (3.5s)')
ax.fill_between(date_stats.index, 3.5, date_stats['mean'], where=date_stats['mean'] > 3.5, 
                alpha=0.1, color=colors['target'])
ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_xlabel('')
ax.set_title('Average Cycle Time Over Collection Period', fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, alpha=0.3)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('/home/claude/chart1_timeline.png', dpi=200, bbox_inches='tight')
plt.close()

# Chart 2: Operator performance comparison
fig, ax = plt.subplots(figsize=(9, 5))
pilot_stats = main.groupby('pilot')['implied_cycle'].agg(['mean', 'median', 'std', 'count'])
pilot_stats = pilot_stats.sort_values('mean')
bars = ax.bar(pilot_stats.index, pilot_stats['mean'], color=[colors['accent'] if m < 3.5 else colors['secondary'] for m in pilot_stats['mean']], 
              edgecolor='white', linewidth=1.5, width=0.6)
ax.errorbar(pilot_stats.index, pilot_stats['mean'], yerr=pilot_stats['std']/np.sqrt(pilot_stats['count']),
            fmt='none', color=colors['text'], capsize=5, alpha=0.7)
ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8, label='Target')
for i, (idx, row) in enumerate(pilot_stats.iterrows()):
    ax.text(i, row['mean'] + 0.08, f'{row["mean"]:.2f}s', ha='center', fontsize=11, fontweight='bold', color=colors['text'])
    ax.text(i, row['mean'] - 0.15, f'n={int(row["count"])}', ha='center', fontsize=9, color='white')
ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Teleoperator Performance Comparison', fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/claude/chart2_operators.png', dpi=200, bbox_inches='tight')
plt.close()

# Chart 3: Config performance
fig, ax = plt.subplots(figsize=(10, 5))
config_stats = main.groupby('config')['implied_cycle'].agg(['mean', 'count'])
config_stats = config_stats[config_stats['count'] >= 50].sort_values('mean')
bar_colors = [colors['accent'] if m < 3.5 else colors['secondary'] for m in config_stats['mean']]
bars = ax.barh(config_stats.index, config_stats['mean'], color=bar_colors, edgecolor='white', height=0.6)
ax.axvline(x=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8, label='Target')
for i, (idx, row) in enumerate(config_stats.iterrows()):
    ax.text(row['mean'] + 0.05, i, f'{row["mean"]:.2f}s (n={int(row["count"])})', va='center', fontsize=9, color=colors['text'])
ax.set_xlabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Performance by Robot Configuration', fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/claude/chart3_configs.png', dpi=200, bbox_inches='tight')
plt.close()

# Chart 4: Distribution histogram
fig, ax = plt.subplots(figsize=(9, 5))
ax.hist(main['implied_cycle'], bins=50, color=colors['secondary'], edgecolor='white', alpha=0.85)
ax.axvline(x=3.5, color=colors['target'], linestyle='--', linewidth=2.5, label='Target (3.5s)')
ax.axvline(x=main['implied_cycle'].mean(), color=colors['primary'], linestyle='-', linewidth=2, label=f'Mean ({main["implied_cycle"].mean():.2f}s)')
pct_under = (main['implied_cycle'] < 3.5).mean() * 100
ax.text(3.5, ax.get_ylim()[1]*0.9, f'  {pct_under:.0f}% under target', fontsize=11, color=colors['target'], fontweight='bold')
ax.set_xlabel('Implied Cycle Time per Package (sec)', fontsize=12, color=colors['text'])
ax.set_ylabel('Frequency', fontsize=12, color=colors['text'])
ax.set_title('Distribution of Cycle Times', fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/claude/chart4_distribution.png', dpi=200, bbox_inches='tight')
plt.close()

# Chart 5: Speed mode impact
fig, ax = plt.subplots(figsize=(7, 4.5))
speed_data = main.groupby('speed')['implied_cycle'].agg(['mean', 'std', 'count'])
speed_data.index = ['Normal', 'Speed Mode']
bars = ax.bar(speed_data.index, speed_data['mean'], color=[colors['secondary'], colors['accent']], 
              edgecolor='white', width=0.5)
for i, (idx, row) in enumerate(speed_data.iterrows()):
    ax.text(i, row['mean'] + 0.05, f'{row["mean"]:.2f}s', ha='center', fontsize=12, fontweight='bold', color=colors['text'])
ax.axhline(y=3.5, color=colors['target'], linestyle='--', linewidth=2, alpha=0.8, label='Target')
ax.set_ylabel('Avg Cycle Time (sec)', fontsize=12, color=colors['text'])
ax.set_title('Impact of Speed Mode', fontsize=14, fontweight='bold', color=colors['primary'])
ax.legend(fontsize=10)
ax.set_facecolor(colors['bg'])
fig.patch.set_facecolor('white')
ax.grid(True, axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('/home/claude/chart5_speed.png', dpi=200, bbox_inches='tight')
plt.close()

print("All charts saved successfully!")