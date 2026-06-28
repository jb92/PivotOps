/**
 * PivotOps — Python Chart & Analytics Scripts
 * Pre-built Python scripts for visualization and analysis.
 * These run inside Pyodide (browser-side WebAssembly).
 */

// ── Core Charts ───────────────────────────────────────────────────────

export const CHART_STATE_DISTRIBUTION = `
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json

df = pd.DataFrame(__data__['items'])

# PivotOps dark theme
sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

# Color palette matching ADO states
state_colors = {
    'New': '#89b4fa',
    'Active': '#fab387',
    'Resolved': '#a6e3a1',
    'Closed': '#74c7ec',
    'Removed': '#f38ba8',
}

fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Pie chart
state_counts = df['State'].value_counts()
colors = [state_colors.get(s, '#cba6f7') for s in state_counts.index]
axes[0].pie(state_counts.values, labels=state_counts.index, colors=colors,
            autopct='%1.1f%%', startangle=90, textprops={'color': '#cdd6f4'})
axes[0].set_title('Work Items by State', fontsize=14, fontweight='bold', color='#cdd6f4')

# Bar chart by type
type_state = pd.crosstab(df['WorkItemType'], df['State'])
type_state.plot(kind='barh', stacked=True, ax=axes[1],
                color=[state_colors.get(c, '#cba6f7') for c in type_state.columns])
axes[1].set_title('State by Work Item Type', fontsize=14, fontweight='bold', color='#cdd6f4')
axes[1].set_xlabel('Count')
axes[1].legend(loc='lower right', facecolor='#313244', edgecolor='#45475a')

plt.tight_layout()
`;

export const CHART_VELOCITY = `
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json

df = pd.DataFrame(__data__['items'])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

# Group by iteration and sum story points
if 'StoryPoints' in df.columns and 'IterationPath' in df.columns:
    df['StoryPoints'] = pd.to_numeric(df['StoryPoints'], errors='coerce').fillna(0)
    # Only closed items contribute to velocity
    closed = df[df['State'].isin(['Closed', 'Resolved', 'Done'])]
    velocity = closed.groupby('IterationPath')['StoryPoints'].sum().reset_index()
    velocity = velocity.sort_values('IterationPath').tail(10)  # Last 10 sprints

    fig, ax = plt.subplots(figsize=(12, 6))

    bars = ax.bar(range(len(velocity)), velocity['StoryPoints'],
                  color='#89b4fa', edgecolor='#45475a', linewidth=0.5)

    # Add trend line
    if len(velocity) > 2:
        import numpy as np
        x = np.arange(len(velocity))
        y = velocity['StoryPoints'].astype(float).values
        z = np.polyfit(x, y, 1)
        p = np.poly1d(z)
        ax.plot(x, p(x), '--', color='#f38ba8', linewidth=2, label=f'Trend ({z[0]:+.1f}/sprint)')
        ax.legend(facecolor='#313244', edgecolor='#45475a')

    # Average line
    avg = velocity['StoryPoints'].mean()
    ax.axhline(y=avg, color='#a6e3a1', linestyle=':', linewidth=1.5, label=f'Avg: {avg:.1f}')

    ax.set_xticks(range(len(velocity)))
    ax.set_xticklabels([p.split('\\\\')[-1] for p in velocity['IterationPath']], rotation=45, ha='right')
    ax.set_ylabel('Story Points Completed')
    ax.set_title('Sprint Velocity', fontsize=16, fontweight='bold', color='#cdd6f4')
    ax.legend(facecolor='#313244', edgecolor='#45475a')

    plt.tight_layout()
else:
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.text(0.5, 0.5, 'No Story Points or Iteration data available',
            ha='center', va='center', fontsize=14, color='#a6adc8',
            transform=ax.transAxes)
    ax.set_facecolor('#1e1e2e')
`;

export const CHART_SCOPE_DRIFT = `
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import json

df = pd.DataFrame(__data__['items'])
iterations = __data__.get('iterations', [])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

if 'IterationPath' in df.columns and len(iterations) > 0:
    df['CreatedDate'] = pd.to_datetime(df['CreatedDate'], format='ISO8601', utc=True).dt.tz_localize(None)

    # Build iteration lookup: path -> start/end dates
    iter_lookup = {}
    for it in iterations:
        start = it.get('startDate') or it.get('start')
        end = it.get('finishDate') or it.get('finish')
        path = it.get('path', it.get('name', ''))
        if start and end and path:
            iter_lookup[path] = {
                'start': pd.to_datetime(start, utc=True).tz_localize(None),
                'end': pd.to_datetime(end, utc=True).tz_localize(None),
                'name': path.split('\\\\')[-1],
            }

    # Match work items to iterations and classify
    records = []
    for path, info in iter_lookup.items():
        sprint_items = df[df['IterationPath'] == path]
        if len(sprint_items) == 0:
            continue
        planned = len(sprint_items[sprint_items['CreatedDate'] < info['start']])
        added = len(sprint_items[sprint_items['CreatedDate'] >= info['start']])
        total = len(sprint_items)
        completed = len(sprint_items[sprint_items['State'].isin(['Closed', 'Resolved', 'Done'])])
        carryover = total - completed
        drift_pct = ((added / planned * 100) if planned > 0 else 0)
        records.append({
            'Sprint': info['name'],
            'Start': info['start'],
            'Planned': planned,
            'Added': added,
            'Completed': completed,
            'Carryover': carryover,
            'DriftPct': drift_pct,
        })

    if len(records) > 0:
        result = pd.DataFrame(records).sort_values('Start').tail(10)
        x = range(len(result))

        fig, ax1 = plt.subplots(figsize=(14, 7))

        bar_width = 0.22
        x_arr = [i for i in x]
        x_planned = [i - bar_width for i in x_arr]
        x_added = [i for i in x_arr]
        x_completed = [i + bar_width for i in x_arr]

        ax1.bar(x_planned, result['Planned'], width=bar_width, color='#89b4fa',
                label='Planned', edgecolor='#45475a', linewidth=0.5)
        ax1.bar(x_added, result['Added'], width=bar_width, color='#fab387',
                label='Added Mid-Sprint', edgecolor='#45475a', linewidth=0.5)
        ax1.bar(x_completed, result['Completed'], width=bar_width, color='#a6e3a1',
                label='Completed', edgecolor='#45475a', linewidth=0.5)

        ax1.set_xticks(list(x))
        ax1.set_xticklabels(result['Sprint'], rotation=45, ha='right')
        ax1.set_ylabel('Work Items')
        ax1.set_title('Plan vs. Actual — Scope Drift per Sprint', fontsize=16, fontweight='bold', color='#cdd6f4')

        # Drift % line on secondary axis
        ax2 = ax1.twinx()
        ax2.plot(list(x), result['DriftPct'].values, 'D--', color='#f38ba8',
                 linewidth=2, markersize=6, label='Scope Drift %')
        ax2.set_ylabel('Scope Drift %', color='#f38ba8')
        ax2.tick_params(axis='y', labelcolor='#f38ba8')
        ax2.yaxis.set_major_formatter(mticker.PercentFormatter())

        # Merge legends from both axes
        h1, l1 = ax1.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax1.legend(h1 + h2, l1 + l2, loc='upper left', facecolor='#313244', edgecolor='#45475a')

        plt.tight_layout()
    else:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No matching iterations found for current work items',
                ha='center', va='center', fontsize=14, color='#a6adc8',
                transform=ax.transAxes)
        ax.set_facecolor('#1e1e2e')
else:
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.text(0.5, 0.5, 'No Iteration data available\\nSync work items and ensure iterations are configured.',
            ha='center', va='center', fontsize=14, color='#a6adc8',
            transform=ax.transAxes)
    ax.set_facecolor('#1e1e2e')
`;

// ── Advanced Charts ────────────────────────────────────────────────────────

export const CHART_CYCLE_TIME = `
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import json

df = pd.DataFrame(__data__['items'])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

fig, axes = plt.subplots(1, 2, figsize=(14, 6))

if 'CycleTimeDays' in df.columns:
    cycle_data = df.dropna(subset=['CycleTimeDays'])

    if len(cycle_data) > 0:
        # Histogram with KDE
        sns.histplot(data=cycle_data, x='CycleTimeDays', kde=True,
                     ax=axes[0], color='#89b4fa', edgecolor='#45475a')
        
        # Percentile lines
        p50 = cycle_data['CycleTimeDays'].quantile(0.5)
        p85 = cycle_data['CycleTimeDays'].quantile(0.85)
        p95 = cycle_data['CycleTimeDays'].quantile(0.95)
        
        axes[0].axvline(p50, color='#a6e3a1', linestyle='--', label=f'P50: {p50:.1f}d')
        axes[0].axvline(p85, color='#fab387', linestyle='--', label=f'P85: {p85:.1f}d')
        axes[0].axvline(p95, color='#f38ba8', linestyle='--', label=f'P95: {p95:.1f}d')
        
        axes[0].set_title('Cycle Time Distribution', fontsize=14, fontweight='bold')
        axes[0].set_xlabel('Days')
        axes[0].legend(facecolor='#313244', edgecolor='#45475a')

        # Cycle time by work item type (violin plot)
        types_with_data = cycle_data.groupby('WorkItemType').filter(lambda x: len(x) >= 3)
        if len(types_with_data) > 0:
            sns.violinplot(data=types_with_data, x='WorkItemType', y='CycleTimeDays',
                          ax=axes[1], palette='muted', inner='quartile')
            axes[1].set_title('Cycle Time by Type', fontsize=14, fontweight='bold')
            axes[1].set_xlabel('')
            axes[1].set_ylabel('Days')
        else:
            axes[1].text(0.5, 0.5, 'Not enough data for violin plot',
                        ha='center', va='center', color='#a6adc8', transform=axes[1].transAxes)
else:
    for ax in axes:
        ax.text(0.5, 0.5, 'No cycle time data available\\n(Use Analytics endpoint)',
                ha='center', va='center', fontsize=12, color='#a6adc8', transform=ax.transAxes)

plt.tight_layout()
`;

export const CHART_CUMULATIVE_FLOW = `
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json

df = pd.DataFrame(__data__['items'])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

fig, ax = plt.subplots(figsize=(14, 7))

# Parse dates
df['ChangedDate'] = pd.to_datetime(df['ChangedDate'], format='ISO8601', utc=True).dt.tz_localize(None)
df['Week'] = df['ChangedDate'].dt.to_period('W').dt.start_time

# State order (from new to done)
state_order = ['New', 'Active', 'Resolved', 'Closed']
existing_states = [s for s in state_order if s in df['State'].unique()]

# Cumulative count by state and week
cfd_data = []
for week in sorted(df['Week'].unique()):
    mask = df['ChangedDate'] <= week + pd.Timedelta(days=7)
    for state in existing_states:
        count = len(df[mask & (df['State'] == state)])
        cfd_data.append({'Week': week, 'State': state, 'Count': count})

cfd = pd.DataFrame(cfd_data)
pivot = cfd.pivot(index='Week', columns='State', values='Count').fillna(0)
pivot = pivot.reindex(columns=existing_states)

colors = {
    'New': '#89b4fa',
    'Active': '#fab387',
    'Resolved': '#a6e3a1',
    'Closed': '#74c7ec',
}

ax.stackplot(pivot.index, *[pivot[col] for col in pivot.columns],
             labels=pivot.columns,
             colors=[colors.get(c, '#cba6f7') for c in pivot.columns],
             alpha=0.85)

ax.set_title('Cumulative Flow Diagram', fontsize=16, fontweight='bold', color='#cdd6f4')
ax.set_xlabel('Week')
ax.set_ylabel('Work Items')
ax.legend(loc='upper left', facecolor='#313244', edgecolor='#45475a')

plt.xticks(rotation=45, ha='right')
plt.tight_layout()
`;

export const CHART_BURNDOWN = `
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import json

df = pd.DataFrame(__data__['items'])

plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

fig, ax = plt.subplots(figsize=(12, 6))

if 'StoryPoints' in df.columns:
    df['ClosedDate'] = pd.to_datetime(df.get('ClosedDate', pd.NaT), format='ISO8601', utc=True).dt.tz_localize(None)
    df['CreatedDate'] = pd.to_datetime(df['CreatedDate'], format='ISO8601', utc=True).dt.tz_localize(None)
    
    total_points = df['StoryPoints'].sum()
    
    # Daily burndown
    closed = df.dropna(subset=['ClosedDate']).copy()
    closed['Day'] = closed['ClosedDate'].dt.date
    daily_closed = closed.groupby('Day')['StoryPoints'].sum().cumsum()
    
    remaining = total_points - daily_closed
    
    ax.plot(remaining.index, remaining.values, 'o-', color='#89b4fa', linewidth=2, markersize=4, label='Actual')
    
    # Ideal burndown line
    if len(remaining) > 1:
        start = remaining.index[0]
        end = remaining.index[-1]
        ideal_x = [start, end]
        ideal_y = [total_points, 0]
        ax.plot(ideal_x, ideal_y, '--', color='#a6adc8', linewidth=1.5, label='Ideal')
    
    ax.set_title('Sprint Burndown', fontsize=16, fontweight='bold', color='#cdd6f4')
    ax.set_xlabel('Date')
    ax.set_ylabel('Story Points Remaining')
    ax.legend(facecolor='#313244', edgecolor='#45475a')
    plt.xticks(rotation=45, ha='right')
else:
    ax.text(0.5, 0.5, 'No Story Points data available',
            ha='center', va='center', fontsize=14, color='#a6adc8', transform=ax.transAxes)

plt.tight_layout()
`;

export const CHART_MEMBER_VELOCITY = `
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import json

df = pd.DataFrame(__data__['items'])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

if 'StoryPoints' in df.columns and 'AssignedTo' in df.columns:
    df['StoryPoints'] = pd.to_numeric(df['StoryPoints'], errors='coerce').fillna(0)
    closed = df[df['State'].isin(['Closed', 'Resolved', 'Done'])]

    if len(closed) > 0 and closed['StoryPoints'].sum() > 0:
        fig, axes = plt.subplots(1, 2, figsize=(16, 7))

        # Left: total story points per member (top 15)
        member_total = closed.groupby('AssignedTo')['StoryPoints'].sum().sort_values(ascending=True).tail(15)
        colors = sns.color_palette('muted', len(member_total))
        axes[0].barh(range(len(member_total)), member_total.values, color=colors, edgecolor='#45475a', linewidth=0.5)
        axes[0].set_yticks(range(len(member_total)))
        axes[0].set_yticklabels([name[:25] for name in member_total.index], fontsize=9)
        axes[0].set_xlabel('Story Points Completed')
        axes[0].set_title('Total Velocity by Member', fontsize=14, fontweight='bold', color='#cdd6f4')

        # Right: velocity per sprint per member (heatmap, top 10 members x last 8 sprints)
        if 'IterationPath' in df.columns:
            top_members = closed.groupby('AssignedTo')['StoryPoints'].sum().nlargest(10).index
            top_data = closed[closed['AssignedTo'].isin(top_members)]
            pivot = top_data.pivot_table(index='AssignedTo', columns='IterationPath',
                                         values='StoryPoints', aggfunc='sum', fill_value=0)
            # Keep last 8 sprints
            pivot = pivot[sorted(pivot.columns)[-8:]]
            pivot.index = [name[:20] for name in pivot.index]
            sprint_labels = [col.split('\\\\')[-1] for col in pivot.columns]

            sns.heatmap(pivot, annot=True, fmt='.0f', cmap='YlOrRd', ax=axes[1],
                        linewidths=0.5, linecolor='#45475a',
                        xticklabels=sprint_labels,
                        cbar_kws={'label': 'Story Points'})
            axes[1].set_title('Velocity Heatmap (Member x Sprint)', fontsize=14, fontweight='bold', color='#cdd6f4')
            axes[1].set_ylabel('')
            axes[1].tick_params(axis='x', rotation=45)
        else:
            axes[1].text(0.5, 0.5, 'No Iteration data for heatmap',
                        ha='center', va='center', fontsize=12, color='#a6adc8',
                        transform=axes[1].transAxes)

        plt.tight_layout()
    else:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No completed items with Story Points found',
                ha='center', va='center', fontsize=14, color='#a6adc8', transform=ax.transAxes)
        ax.set_facecolor('#1e1e2e')
else:
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.text(0.5, 0.5, 'No Story Points or AssignedTo data available',
            ha='center', va='center', fontsize=14, color='#a6adc8', transform=ax.transAxes)
    ax.set_facecolor('#1e1e2e')
`;

export const CHART_ESTIMATION_ACCURACY = `
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import numpy as np
import json

df = pd.DataFrame(__data__['items'])

sns.set_theme(style="darkgrid")
plt.rcParams.update({
    'figure.facecolor': '#1e1e2e',
    'axes.facecolor': '#2a2a3e',
    'text.color': '#cdd6f4',
    'axes.labelcolor': '#cdd6f4',
    'xtick.color': '#a6adc8',
    'ytick.color': '#a6adc8',
    'grid.color': '#45475a',
})

has_points = 'StoryPoints' in df.columns
has_cycle = 'CycleTimeDays' in df.columns
has_assigned = 'AssignedTo' in df.columns

if has_points and has_cycle and has_assigned:
    df['StoryPoints'] = pd.to_numeric(df['StoryPoints'], errors='coerce')
    df['CycleTimeDays'] = pd.to_numeric(df['CycleTimeDays'], errors='coerce')
    closed = df[df['State'].isin(['Closed', 'Resolved', 'Done'])].dropna(subset=['StoryPoints', 'CycleTimeDays'])
    closed = closed[closed['StoryPoints'] > 0]

    if len(closed) >= 3:
        # Compute expected cycle time per story point (team baseline)
        baseline_days_per_point = closed['CycleTimeDays'].median() / closed['StoryPoints'].median()

        # Per member: expected days (points * baseline) vs actual days
        member_stats = []
        for member, group in closed.groupby('AssignedTo'):
            if len(group) < 2:
                continue
            expected = group['StoryPoints'] * baseline_days_per_point
            actual = group['CycleTimeDays']
            ratio = (actual / expected).median()  # <1 = faster, >1 = slower
            accuracy_pct = (1 - abs(1 - ratio)) * 100  # 100% = perfect match
            items_count = len(group)
            member_stats.append({
                'Member': str(member)[:25],
                'Items': items_count,
                'MedianRatio': ratio,
                'AccuracyPct': accuracy_pct,
                'AvgExpected': expected.mean(),
                'AvgActual': actual.mean(),
            })

        if len(member_stats) >= 2:
            stats = pd.DataFrame(member_stats).sort_values('AccuracyPct', ascending=True)

            fig, axes = plt.subplots(1, 2, figsize=(16, 7))

            # Left: Estimation accuracy leaderboard (horizontal bar)
            colors = []
            for _, row in stats.iterrows():
                if row['MedianRatio'] < 0.85:
                    colors.append('#a6e3a1')   # faster than expected
                elif row['MedianRatio'] > 1.15:
                    colors.append('#f38ba8')   # slower than expected
                else:
                    colors.append('#89b4fa')   # on target

            axes[0].barh(range(len(stats)), stats['AccuracyPct'], color=colors,
                        edgecolor='#45475a', linewidth=0.5)
            axes[0].set_yticks(range(len(stats)))
            axes[0].set_yticklabels(stats['Member'], fontsize=9)
            axes[0].set_xlabel('Estimation Accuracy %')
            axes[0].set_title('Estimation Accuracy Leaderboard', fontsize=14, fontweight='bold', color='#cdd6f4')
            axes[0].axvline(x=100, color='#a6adc8', linestyle=':', linewidth=1, alpha=0.5)

            # Legend for colors
            from matplotlib.patches import Patch
            legend_elements = [
                Patch(facecolor='#a6e3a1', edgecolor='#45475a', label='Faster than estimate'),
                Patch(facecolor='#89b4fa', edgecolor='#45475a', label='On target (\\u00b115%)'),
                Patch(facecolor='#f38ba8', edgecolor='#45475a', label='Slower than estimate'),
            ]
            axes[0].legend(handles=legend_elements, loc='lower right',
                          facecolor='#313244', edgecolor='#45475a', fontsize=8)

            # Right: Expected vs Actual (scatter with diagonal)
            axes[1].scatter(stats['AvgExpected'], stats['AvgActual'],
                           s=stats['Items'] * 20, c='#cba6f7', alpha=0.7,
                           edgecolors='#45475a', linewidths=0.5)

            # Perfect estimation line
            max_val = max(stats['AvgExpected'].max(), stats['AvgActual'].max()) * 1.1
            axes[1].plot([0, max_val], [0, max_val], '--', color='#a6adc8', linewidth=1.5, label='Perfect estimate')

            # Label each point
            for _, row in stats.iterrows():
                axes[1].annotate(row['Member'], (row['AvgExpected'], row['AvgActual']),
                                fontsize=7, color='#cdd6f4', ha='left', va='bottom')

            axes[1].set_xlabel('Avg Expected Days (points x baseline)')
            axes[1].set_ylabel('Avg Actual Days (cycle time)')
            axes[1].set_title('Expected vs Actual Duration', fontsize=14, fontweight='bold', color='#cdd6f4')
            axes[1].legend(facecolor='#313244', edgecolor='#45475a')

            plt.tight_layout()
        else:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'Need at least 2 members with completed items',
                    ha='center', va='center', fontsize=14, color='#a6adc8', transform=ax.transAxes)
            ax.set_facecolor('#1e1e2e')
    else:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'Not enough completed items with Story Points and Cycle Time',
                ha='center', va='center', fontsize=14, color='#a6adc8', transform=ax.transAxes)
        ax.set_facecolor('#1e1e2e')
else:
    fig, ax = plt.subplots(figsize=(12, 6))
    missing = [f for f, h in [('StoryPoints', has_points), ('CycleTimeDays', has_cycle), ('AssignedTo', has_assigned)] if not h]
    ax.text(0.5, 0.5, f'Missing data: {", ".join(missing)}\\nUse Analytics endpoint for CycleTimeDays.',
            ha='center', va='center', fontsize=12, color='#a6adc8', transform=ax.transAxes)
    ax.set_facecolor('#1e1e2e')
`;

// ── Predictive Analytics ────────────────────────────────────────

export const ANALYSIS_SPRINT_PREDICTION = `
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import GradientBoostingRegressor
import json

df = pd.DataFrame(__data__['items'])
__result__ = {}

# Sanitize numeric columns
for col in ['StoryPoints', 'CycleTimeDays', 'LeadTimeDays']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

if 'StoryPoints' in df.columns and 'IterationPath' in df.columns:
    closed = df[df['State'].isin(['Closed', 'Resolved', 'Done'])]
    velocity = closed.groupby('IterationPath')['StoryPoints'].sum().reset_index()
    velocity = velocity.sort_values('IterationPath')
    
    if len(velocity) >= 3:
        X = np.arange(len(velocity)).reshape(-1, 1)
        y = velocity['StoryPoints'].astype(float).values
        
        # Linear regression for trend
        lr = LinearRegression()
        lr.fit(X, y)
        
        # Predict next 3 sprints
        future_X = np.arange(len(velocity), len(velocity) + 3).reshape(-1, 1)
        predictions = lr.predict(future_X)
        
        # Confidence interval (based on residual std)
        residuals = y - lr.predict(X)
        std = np.std(residuals)
        
        __result__ = {
            'predictions': [
                {
                    'sprint': f'Sprint +{i+1}',
                    'predicted_velocity': round(max(0, float(predictions[i])), 1),
                    'confidence_low': round(max(0, float(predictions[i] - 1.96 * std)), 1),
                    'confidence_high': round(float(predictions[i] + 1.96 * std), 1),
                }
                for i in range(3)
            ],
            'trend': round(float(lr.coef_[0]), 2),
            'trend_direction': 'improving' if lr.coef_[0] > 0.5 else ('declining' if lr.coef_[0] < -0.5 else 'stable'),
            'average_velocity': round(float(np.mean(y)), 1),
            'velocity_std': round(float(std), 1),
            'r_squared': round(float(lr.score(X, y)), 3),
            'historical': [
                {'sprint': str(velocity.iloc[i]['IterationPath']).split('\\\\')[-1],
                 'velocity': float(velocity.iloc[i]['StoryPoints'])}
                for i in range(len(velocity))
            ],
        }
    else:
        __result__ = {'error': 'Need at least 3 sprints of data for predictions.'}
else:
    __result__ = {'error': 'No StoryPoints or IterationPath data available.'}
`;

export const ANALYSIS_COMPLETION_FORECAST = `
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import json

df = pd.DataFrame(__data__['items'])
backlog_points = __data__.get('backlog_points', 0)
__result__ = {}

# Sanitize numeric columns
for col in ['StoryPoints', 'CycleTimeDays', 'LeadTimeDays']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

if 'StoryPoints' in df.columns and 'IterationPath' in df.columns:
    closed = df[df['State'].isin(['Closed', 'Resolved', 'Done'])]
    velocity = closed.groupby('IterationPath')['StoryPoints'].sum().reset_index()
    velocity = velocity.sort_values('IterationPath')
    
    if len(velocity) >= 3 and backlog_points > 0:
        avg_velocity = velocity['StoryPoints'].mean()
        std_velocity = velocity['StoryPoints'].std()
        
        # Monte Carlo simulation (1000 runs)
        np.random.seed(42)
        simulations = 1000
        sprints_to_complete = []
        
        for _ in range(simulations):
            remaining = backlog_points
            sprints = 0
            while remaining > 0 and sprints < 100:
                v = max(1, np.random.normal(avg_velocity, std_velocity))
                remaining -= v
                sprints += 1
            sprints_to_complete.append(sprints)
        
        sprints_array = np.array(sprints_to_complete)
        
        __result__ = {
            'backlog_points': backlog_points,
            'average_velocity': round(float(avg_velocity), 1),
            'p50_sprints': int(np.percentile(sprints_array, 50)),
            'p85_sprints': int(np.percentile(sprints_array, 85)),
            'p95_sprints': int(np.percentile(sprints_array, 95)),
            'simulations': simulations,
        }
    else:
        __result__ = {'error': 'Need at least 3 sprints of velocity data and a backlog estimate.'}
else:
    __result__ = {'error': 'No StoryPoints or IterationPath data available.'}
`;

// ── Analytics summary ──────────────────────────────────────────

export const ANALYSIS_SUMMARY = `
import pandas as pd
import json

df = pd.DataFrame(__data__['items'])
__result__ = {}

# Sanitize numeric columns
for col in ['StoryPoints', 'CycleTimeDays', 'LeadTimeDays']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

total = len(df)
by_state = df['State'].value_counts().to_dict()
by_type = df['WorkItemType'].value_counts().to_dict() if 'WorkItemType' in df.columns else {}

summary = {
    'total_items': total,
    'by_state': by_state,
    'by_type': by_type,
}

if 'StoryPoints' in df.columns:
    summary['total_points'] = float(df['StoryPoints'].sum())
    summary['avg_points'] = round(float(df['StoryPoints'].mean()), 1)

if 'CycleTimeDays' in df.columns:
    ct = df['CycleTimeDays'].dropna()
    if len(ct) > 0:
        summary['cycle_time'] = {
            'median': round(float(ct.median()), 1),
            'p85': round(float(ct.quantile(0.85)), 1),
            'mean': round(float(ct.mean()), 1),
        }

if 'AssignedTo' in df.columns:
    top_assignees = df['AssignedTo'].value_counts().head(10).to_dict()
    summary['top_assignees'] = top_assignees

__result__ = summary
`;
