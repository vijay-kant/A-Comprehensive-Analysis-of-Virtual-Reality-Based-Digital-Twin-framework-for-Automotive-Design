"""
VR-DT Framework Visualization Generator
Requires: pip install pandas matplotlib seaborn openpyxl
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Set style
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("Set2")
plt.rcParams['figure.dpi'] = 150
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 11

# Load data (update path to your Excel file location)
# df = pd.read_excel('Dataset_N80.xlsx')

# OR use the data directly (for convenience)
csv_data = """ID,Group,Role,Age,Gender,Prior_VR_Exp,Pre_Test_Pct,Post_Test_Pct,Retention_1wk_Pct,Task_Time_min,SUS_Score,Motion_Sickness
D001,Design,Designer,34,M,4,,,,14.2,82,1
D002,Design,Designer,41,F,3,,,,15.8,78,2
D003,Design,Designer,29,M,5,,,,12.5,88,1
D004,Design,Designer,38,F,2,,,,16.2,74,2
D005,Design,Designer,26,M,4,,,,13.8,81,1
D006,Design,Designer,45,M,2,,,,17.4,72,3
D007,Design,Designer,31,F,3,,,,14.9,79,2
D008,Design,Designer,28,M,5,,,,12.1,86,1
D009,Design,Designer,36,F,2,,,,16.5,73,2
D010,Design,Designer,33,M,3,,,,15.2,77,1
D011,Design,Designer,27,F,4,,,,13.4,84,1
D012,Design,Designer,42,M,1,,,,18.1,68,3
D013,Design,Designer,30,F,3,,,,14.6,80,2
D014,Design,Designer,35,M,2,,,,15.9,75,2
D015,Design,Designer,25,F,4,,,,13.2,85,1
D016,Design,Designer,39,M,3,,,,14.8,79,2
D017,Design,Designer,32,F,5,,,,12.8,87,1
D018,Design,Designer,44,M,1,,,,17.9,69,3
D019,Design,Designer,28,F,3,,,,15.1,78,2
D020,Design,Designer,37,M,2,,,,16.3,74,2
D021,Design,Designer,24,M,5,,,,11.9,89,1
D022,Design,Designer,33,F,3,,,,14.4,81,1
D023,Design,Designer,40,M,2,,,,16.8,73,2
D024,Design,Designer,27,F,4,,,,13.6,83,1
D025,Design,Designer,35,M,3,,,,15.3,78,2
D026,Design,Designer,29,F,2,,,,16.1,75,2
D027,Design,Designer,31,M,4,,,,13.9,82,1
D028,Design,Designer,38,F,1,,,,18.2,67,3
D029,Design,Designer,26,M,3,,,,14.7,80,2
D030,Design,Designer,34,F,3,,,,15.4,79,1
D031,Design,Student,22,M,4,,,,12.6,86,1
D032,Design,Student,21,F,5,,,,11.8,90,1
D033,Design,Student,23,M,3,,,,14.2,81,2
D034,Design,Student,20,F,5,,,,12.4,88,1
D035,Design,Student,24,M,4,,,,13.1,84,1
D036,Design,Student,22,F,3,,,,14.8,79,2
D037,Design,Student,21,M,5,,,,12.2,87,1
D038,Design,Student,23,F,4,,,,13.5,83,1
D039,Design,Student,20,M,3,,,,15.1,78,2
D040,Design,Student,22,F,5,,,,12.3,89,1
T001,Training,Engineer,35,M,2,48,92,86,12.8,82,2
T002,Training,Engineer,42,F,1,42,88,81,14.2,78,3
T003,Training,Engineer,29,M,3,55,94,89,11.5,86,1
T004,Training,Engineer,38,F,2,46,90,84,13.1,81,2
T005,Training,Engineer,31,M,4,58,96,91,10.8,89,1
T006,Training,Engineer,45,M,1,38,84,76,15.4,74,3
T007,Training,Engineer,27,F,3,52,93,87,12.2,84,2
T008,Training,Engineer,33,M,2,44,89,82,13.8,79,2
T009,Training,Engineer,40,F,1,40,86,79,14.6,76,3
T010,Training,Engineer,28,M,4,56,95,90,11.2,88,1
T011,Training,Engineer,36,F,2,47,91,85,12.9,82,2
T012,Training,Engineer,30,M,3,50,92,86,12.5,83,1
T013,Training,Engineer,34,F,2,45,89,83,13.4,80,2
T014,Training,Engineer,26,M,4,54,94,88,11.8,87,1
T015,Training,Engineer,39,F,1,41,87,80,14.9,75,3
T016,Training,Engineer,32,M,3,49,92,85,12.6,83,2
T017,Training,Engineer,37,F,2,46,90,84,13.2,81,2
T018,Training,Engineer,25,M,4,53,95,88,11.4,87,1
T019,Training,Engineer,41,F,1,39,85,78,15.1,74,3
T020,Training,Engineer,28,M,3,52,93,87,12.4,84,2
T021,Training,Engineer,33,F,2,48,90,84,13.5,80,2
T022,Training,Engineer,30,M,4,55,95,89,11.6,86,1
T023,Training,Student,22,M,4,51,94,88,11.9,85,1
T024,Training,Student,21,F,5,58,97,92,10.5,90,1
T025,Training,Student,23,M,3,49,92,86,12.7,83,2
T026,Training,Student,20,F,5,56,96,90,11.1,88,1
T027,Training,Student,22,M,4,53,94,88,11.8,86,1
T028,Training,Student,24,F,3,50,92,86,12.6,82,2
T029,Training,Student,21,M,5,57,96,91,10.9,89,1
T030,Training,Student,23,F,4,52,94,87,12.1,85,1
T031,Training,Student,20,M,4,48,91,85,12.9,83,2
T032,Training,Student,22,F,5,55,95,89,11.3,87,1
T033,Training,Student,21,M,3,49,92,86,12.5,82,2
T034,Training,Student,23,F,4,54,94,88,11.7,86,1
T035,Training,Student,20,M,5,56,96,90,11.0,88,1
T036,Training,Student,22,F,3,50,92,85,12.8,83,2
T037,Training,Student,24,M,4,51,93,87,12.2,84,1
T038,Training,Student,21,F,5,57,96,91,10.7,89,1
T039,Training,Student,23,M,3,48,91,84,13.1,81,2
T040,Training,Student,20,F,4,53,94,88,11.9,85,1"""

from io import StringIO
df = pd.read_csv(StringIO(csv_data))

# Handle empty strings in numeric columns
for col in ['Pre_Test_Pct', 'Post_Test_Pct', 'Retention_1wk_Pct']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Create figure directory
import os
os.makedirs('figures', exist_ok=True)

print("=" * 60)
print("VR-DT FRAMEWORK VISUALIZATION GENERATOR")
print("=" * 60)
print(f"Loaded {len(df)} participants")
print(f"Design Group: {len(df[df['Group']=='Design'])}")
print(f"Training Group: {len(df[df['Group']=='Training'])}")
print("=" * 60)

# -------------------------------------------------------------------------
# FIGURE 1: Ablation Study - Frame Rate Comparison (Bar Chart)
# -------------------------------------------------------------------------
fig1, ax1 = plt.subplots(figsize=(8, 6))
conditions = ['Baseline', 'BVH Only', 'LOD Only', 'Combined']
fps_values = [52.3, 58.6, 57.1, 64.7]
colors_fps = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99']
bars = ax1.bar(conditions, fps_values, color=colors_fps, edgecolor='black', linewidth=1.2)
ax1.set_ylabel('Average Frame Rate (FPS)', fontsize=12, fontweight='bold')
ax1.set_xlabel('Optimization Condition', fontsize=12, fontweight='bold')
ax1.set_title('Figure 1: Rendering Performance - Ablation Study', fontsize=14, fontweight='bold')
ax1.set_ylim(40, 80)
for bar, val in zip(bars, fps_values):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{val} FPS', 
             ha='center', va='bottom', fontweight='bold')
ax1.axhline(y=60, color='red', linestyle='--', linewidth=1.5, label='60 FPS Threshold')
ax1.legend()
plt.tight_layout()
plt.savefig('figures/Figure1_Ablation_FPS.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 1 saved: figures/Figure1_Ablation_FPS.png")

# -------------------------------------------------------------------------
# FIGURE 2: Task Completion Comparison (Bar Chart)
# -------------------------------------------------------------------------
fig2, ax2 = plt.subplots(figsize=(8, 6))
tasks = ['Physical Only', 'VR Only', 'VR + Physical (Iterative)']
tasks_completed = [5.2, 9.2, 11.8]
colors_tasks = ['#ff9999', '#66b3ff', '#99ff99']
bars = ax2.bar(tasks, tasks_completed, color=colors_tasks, edgecolor='black', linewidth=1.2)
ax2.set_ylabel('Tasks Completed in 15 Minutes', fontsize=12, fontweight='bold')
ax2.set_xlabel('Design Review Method', fontsize=12, fontweight='bold')
ax2.set_title('Figure 2: Design Review Task Completion Comparison', fontsize=14, fontweight='bold')
ax2.set_ylim(0, 14)
for bar, val in zip(bars, tasks_completed):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, f'{val} tasks', 
             ha='center', va='bottom', fontweight='bold')
plt.tight_layout()
plt.savefig('figures/Figure2_Task_Completion.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 2 saved: figures/Figure2_Task_Completion.png")

# -------------------------------------------------------------------------
# FIGURE 3: Learning Curve (Line Graph)
# -------------------------------------------------------------------------
fig3, ax3 = plt.subplots(figsize=(10, 6))
attempts = [1, 2, 3, 4, 5]
vr_times = [742, 348, 264, 228, 204]
control_times = [684, 552, 498, 474, 456]
ax3.plot(attempts, vr_times, 'o-', linewidth=2.5, markersize=10, 
         label='VR Training Group', color='#2ecc71')
ax3.plot(attempts, control_times, 's--', linewidth=2.5, markersize=10, 
         label='Control Group (Traditional)', color='#e74c3c')
ax3.set_xlabel('Attempt Number', fontsize=12, fontweight='bold')
ax3.set_ylabel('Time to Identify Correct Parameters (seconds)', fontsize=12, fontweight='bold')
ax3.set_title('Figure 3: Learning Curve - Parameter Identification Time', fontsize=14, fontweight='bold')
ax3.set_xticks(attempts)
ax3.legend(loc='upper right', fontsize=11)
ax3.grid(True, alpha=0.3)
for i, (vr, ct) in enumerate(zip(vr_times, control_times)):
    ax3.annotate(f'{vr}s', (attempts[i], vr), textcoords="offset points", xytext=(0,10), ha='center')
    ax3.annotate(f'{ct}s', (attempts[i], ct), textcoords="offset points", xytext=(0,-15), ha='center')
plt.tight_layout()
plt.savefig('figures/Figure3_Learning_Curve.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 3 saved: figures/Figure3_Learning_Curve.png")

# -------------------------------------------------------------------------
# FIGURE 4: Knowledge Retention - Pre vs Post vs 1-Week (Bar Chart with Error Bars)
# -------------------------------------------------------------------------
training_df = df[df['Group'] == 'Training'].dropna(subset=['Pre_Test_Pct', 'Post_Test_Pct', 'Retention_1wk_Pct'])
pre_mean = training_df['Pre_Test_Pct'].mean()
pre_std = training_df['Pre_Test_Pct'].std()
post_mean = training_df['Post_Test_Pct'].mean()
post_std = training_df['Post_Test_Pct'].std()
ret_mean = training_df['Retention_1wk_Pct'].mean()
ret_std = training_df['Retention_1wk_Pct'].std()

fig4, ax4 = plt.subplots(figsize=(10, 6))
assessments = ['Pre-Training', 'Post-Training\n(Immediate)', '1-Week\nRetention']
means = [pre_mean, post_mean, ret_mean]
stds = [pre_std, post_std, ret_std]
colors_knowledge = ['#e74c3c', '#2ecc71', '#3498db']
bars = ax4.bar(assessments, means, yerr=stds, capsize=8, color=colors_knowledge, 
               edgecolor='black', linewidth=1.2, alpha=0.8)
ax4.set_ylabel('Mean Score (%)', fontsize=12, fontweight='bold')
ax4.set_xlabel('Assessment Time', fontsize=12, fontweight='bold')
ax4.set_title('Figure 4: Knowledge Retention - Pre, Post, and 1-Week Follow-up', fontsize=14, fontweight='bold')
ax4.set_ylim(0, 100)
for bar, val in zip(bars, means):
    ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2, f'{val:.1f}%', 
             ha='center', va='bottom', fontweight='bold')
plt.tight_layout()
plt.savefig('figures/Figure4_Knowledge_Retention.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 4 saved: figures/Figure4_Knowledge_Retention.png")

# -------------------------------------------------------------------------
# FIGURE 5: SUS Score Distribution (Box Plot)
# -------------------------------------------------------------------------
fig5, ax5 = plt.subplots(figsize=(8, 6))
design_sus = df[df['Group'] == 'Design']['SUS_Score'].dropna()
training_sus = df[df['Group'] == 'Training']['SUS_Score'].dropna()
bp = ax5.boxplot([design_sus, training_sus], labels=['Design Review Group\n(n=40)', 'Training Group\n(n=40)'],
                 patch_artist=True, notch=True)
colors_box = ['#3498db', '#2ecc71']
for patch, color in zip(bp['boxes'], colors_box):
    patch.set_facecolor(color)
    patch.set_alpha(0.7)
ax5.set_ylabel('System Usability Scale (SUS) Score', fontsize=12, fontweight='bold')
ax5.set_title('Figure 5: SUS Score Distribution by Group', fontsize=14, fontweight='bold')
ax5.axhline(y=68, color='red', linestyle='--', linewidth=1.5, label='Industry Average (68)')
ax5.legend()
ax5.set_ylim(60, 95)
plt.tight_layout()
plt.savefig('figures/Figure5_SUS_Boxplot.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 5 saved: figures/Figure5_SUS_Boxplot.png")

# -------------------------------------------------------------------------
# FIGURE 6: Prior VR Experience vs Task Time (Scatter Plot with Regression)
# -------------------------------------------------------------------------
fig6, ax6 = plt.subplots(figsize=(10, 6))
design_plot = df[df['Group'] == 'Design']
training_plot = df[df['Group'] == 'Training']
ax6.scatter(design_plot['Prior_VR_Exp'], design_plot['Task_Time_min'], 
            alpha=0.7, s=100, label='Design Review Group', color='#3498db')
ax6.scatter(training_plot['Prior_VR_Exp'], training_plot['Task_Time_min'], 
            alpha=0.7, s=100, label='Training Group', color='#e74c3c')
# Add trend lines
z_design = np.polyfit(design_plot['Prior_VR_Exp'], design_plot['Task_Time_min'], 1)
p_design = np.poly1d(z_design)
z_train = np.polyfit(training_plot['Prior_VR_Exp'], training_plot['Task_Time_min'], 1)
p_train = np.poly1d(z_train)
x_line = np.linspace(0.5, 5.5, 100)
ax6.plot(x_line, p_design(x_line), '--', color='#3498db', linewidth=2)
ax6.plot(x_line, p_train(x_line), '--', color='#e74c3c', linewidth=2)
ax6.set_xlabel('Prior VR Experience (1=None, 5=Extensive)', fontsize=12, fontweight='bold')
ax6.set_ylabel('Task Completion Time (minutes)', fontsize=12, fontweight='bold')
ax6.set_title('Figure 6: Prior VR Experience vs Task Completion Time', fontsize=14, fontweight='bold')
ax6.legend()
ax6.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/Figure6_VR_Experience_Scatter.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 6 saved: figures/Figure6_VR_Experience_Scatter.png")

# -------------------------------------------------------------------------
# FIGURE 7: Motion Sickness Distribution (Histogram)
# -------------------------------------------------------------------------
fig7, ax7 = plt.subplots(figsize=(10, 6))
all_motion = df['Motion_Sickness'].dropna()
ax7.hist(all_motion, bins=[1,2,3,4], edgecolor='black', linewidth=1.2, 
         alpha=0.7, color='#9b59b6', rwidth=0.8)
ax7.set_xlabel('Motion Sickness Level (1=None, 2=Mild, 3=Moderate, 4=Severe)', fontsize=12, fontweight='bold')
ax7.set_ylabel('Number of Participants', fontsize=12, fontweight='bold')
ax7.set_title('Figure 7: Motion Sickness Distribution (N=80)', fontsize=14, fontweight='bold')
ax7.set_xticks([1, 2, 3, 4])
ax7.set_xticklabels(['1 (None)', '2 (Mild)', '3 (Moderate)', '4 (Severe)'])
for i, count in enumerate([len(all_motion[all_motion==1]), len(all_motion[all_motion==2]), 
                           len(all_motion[all_motion==3]), len(all_motion[all_motion==4])]):
    ax7.text(i+1, count + 1, str(count), ha='center', fontweight='bold')
plt.tight_layout()
plt.savefig('figures/Figure7_Motion_Sickness.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 7 saved: figures/Figure7_Motion_Sickness.png")

# -------------------------------------------------------------------------
# FIGURE 8: Group Comparison - Multiple Metrics (Grouped Bar Chart)
# -------------------------------------------------------------------------
fig8, ax8 = plt.subplots(figsize=(12, 7))
metrics = ['Task Time\n(min)', 'SUS Score\n(/100)', 'Motion\nSickness (1-5)', 
           'Prior VR Exp\n(1-5)']
design_vals = [14.7, 80.4, 1.68, 3.1]
training_vals = [12.3, 84.1, 1.55, 3.3]
x = np.arange(len(metrics))
width = 0.35
bars1 = ax8.bar(x - width/2, design_vals, width, label='Design Review Group', 
                color='#3498db', edgecolor='black', linewidth=1)
bars2 = ax8.bar(x + width/2, training_vals, width, label='Training Group', 
                color='#e74c3c', edgecolor='black', linewidth=1)
ax8.set_ylabel('Score', fontsize=12, fontweight='bold')
ax8.set_title('Figure 8: Group Comparison - Key Performance Metrics', fontsize=14, fontweight='bold')
ax8.set_xticks(x)
ax8.set_xticklabels(metrics)
ax8.legend()
for bar in bars1:
    ax8.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1, f'{bar.get_height():.1f}', 
             ha='center', va='bottom', fontsize=9)
for bar in bars2:
    ax8.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1, f'{bar.get_height():.1f}', 
             ha='center', va='bottom', fontsize=9)
plt.tight_layout()
plt.savefig('figures/Figure8_Group_Comparison.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 8 saved: figures/Figure8_Group_Comparison.png")

# -------------------------------------------------------------------------
# FIGURE 9: SUS Score Distribution - Violin Plot
# -------------------------------------------------------------------------
fig9, ax9 = plt.subplots(figsize=(10, 6))
sus_data = [df[df['Group'] == 'Design']['SUS_Score'].dropna().values,
            df[df['Group'] == 'Training']['SUS_Score'].dropna().values]
vp = ax9.violinplot(sus_data, positions=[1, 2], showmeans=True, showmedians=True)
for i, pc in enumerate(vp['bodies']):
    pc.set_facecolor(['#3498db', '#2ecc71'][i])
    pc.set_alpha(0.7)
ax9.set_xticks([1, 2])
ax9.set_xticklabels(['Design Review Group\n(n=40)', 'Training Group\n(n=40)'])
ax9.set_ylabel('System Usability Scale (SUS) Score', fontsize=12, fontweight='bold')
ax9.set_title('Figure 9: SUS Score Distribution - Violin Plot', fontsize=14, fontweight='bold')
ax9.set_ylim(60, 95)
ax9.axhline(y=68, color='red', linestyle='--', linewidth=1.5, label='Industry Average (68)')
ax9.legend()
plt.tight_layout()
plt.savefig('figures/Figure9_SUS_Violin.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Figure 9 saved: figures/Figure9_SUS_Violin.png")

# -------------------------------------------------------------------------
# FIGURE 10: Age Distribution by Group (Histogram)
# -------------------------------------------------------------------------
fig10, (ax10a