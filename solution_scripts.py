import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = 'marketing_data_fb.xlsx'
xlsx = pd.ExcelFile(file_path)

# Load the data from the relevant sheet
df = pd.read_excel(xlsx, 'Revised Cleaned Data')

# Aggregate data by campaign ID
campaign_data = df.groupby('campaign ID').agg({
    'Reach': 'sum',
    'Impressions': 'sum',
    'Frequency': 'mean',
    'Clicks': 'sum',
    'Unique Clicks': 'sum',
    'Unique Link Clicks (ULC)': 'sum',
    'Click-Through Rate (CTR)': 'mean',
    'Unique Click-Through Rate (Unique CTR)': 'mean',
    'Cost Per Click (CPC)': 'mean',
    'Cost per Result (CPR)': 'mean',
    'Amount Spent in INR': 'sum',
}).reset_index()

# Define thresholds and identify underperforming campaigns
campaign_data['Rank_Reach'] = campaign_data['Reach'].rank(ascending=True)
campaign_data['Rank_Clicks'] = campaign_data['Clicks'].rank(ascending=True)
campaign_data['Rank_Impressions'] = campaign_data['Impressions'].rank(ascending=True)
campaign_data['Rank_Frequency'] = campaign_data['Frequency'].rank(ascending=True)
campaign_data['Rank_Unique_Clicks'] = campaign_data['Unique Clicks'].rank(ascending=True)
campaign_data['Rank_ULC'] = campaign_data['Unique Link Clicks (ULC)'].rank(ascending=True)
campaign_data['Rank_CTR'] = campaign_data['Click-Through Rate (CTR)'].rank(ascending=True)
campaign_data['Rank_Unique_CTR'] = campaign_data['Unique Click-Through Rate (Unique CTR)'].rank(ascending=True)
campaign_data['Rank_CPC'] = campaign_data['Cost Per Click (CPC)'].rank(ascending=False)
campaign_data['Rank_CPR'] = campaign_data['Cost per Result (CPR)'].rank(ascending=False)
campaign_data['Rank_Amount_Spent'] = campaign_data['Amount Spent in INR'].rank(ascending=False)

# Summarize rankings to determine overall underperformance
campaign_data['Low_Performance_Score'] = (
    campaign_data[['Rank_CTR', 'Rank_Unique_CTR', 'Rank_Reach', 'Rank_Clicks', 'Rank_Impressions', 'Rank_Frequency', 'Rank_Unique_Clicks', 'Rank_ULC']].sum(axis=1)
)

# Calculate High Cost Score
campaign_data['High_Cost_Score'] = (
    campaign_data[['Rank_CPC', 'Rank_CPR', 'Rank_Amount_Spent']].sum(axis=1)
)

# Identify top 2 low-performing and high-cost campaigns
low_performance_campaigns = campaign_data.sort_values(by='Low_Performance_Score').head(2)
high_cost_campaigns = campaign_data.sort_values(by='High_Cost_Score', ascending=False).head(2)

# Set visualization style
sns.set(style="whitegrid")
plt.figure(figsize=(15, 15))

# Plot Performance Score for each campaign
plt.figure(figsize=(12, 7))
palette = ['lightblue' if (x not in low_performance_campaigns['campaign ID'].values) else 'red' for x in campaign_data['campaign ID']]
sns.barplot(x='campaign ID', y='Low_Performance_Score', data=campaign_data, palette=palette)
plt.title('Overall Performance Score by Campaign (Highlighted Low Performers)')
plt.xlabel('Campaign ID')
plt.ylabel('Low Performance Score')
plt.xticks(rotation=45)
plt.show()

# Plot Cost Performance Score for each campaign
plt.figure(figsize=(12, 7))
palette = ['lightblue' if (x not in high_cost_campaigns['campaign ID'].values) else 'orange' for x in campaign_data['campaign ID']]
sns.barplot(x='campaign ID', y='High_Cost_Score', data=campaign_data, palette=palette)
plt.title('Overall Cost Performance Score by Campaign (Highlighted High Cost)')
plt.xlabel('Campaign ID')
plt.ylabel('High Cost Score')
plt.xticks(rotation=45)
plt.show()

# Print campaigns recommended for discontinuation
print("Overall Campaign Ranking:")
print(campaign_data[['campaign ID', 'Low_Performance_Score', 'High_Cost_Score']].sort_values(by=['Low_Performance_Score', 'High_Cost_Score'], ascending=[True, False]))

# Combine the scores for final ranking
campaign_data['Composite_Score'] = campaign_data['Low_Performance_Score'] + campaign_data['High_Cost_Score']

# Rank campaigns by Composite Score
campaign_data['Composite_Rank'] = campaign_data['Composite_Score'].rank(ascending=True)

# Display campaigns ranked by composite score
print("Campaigns Ranked by Combined Low Performance and High Cost Scores:")
print(campaign_data[['campaign ID', 'Composite_Score', 'Composite_Rank']].sort_values(by='Composite_Rank'))

# Identify top 2 low-performing and high-cost campaigns based on Composite Score
top_2_low_performing_high_cost_campaigns = campaign_data.sort_values(by='Composite_Score').head(2)
print("\nTop 2 Low-Performing and High-Cost Campaigns:")
print(top_2_low_performing_high_cost_campaigns[['campaign ID', 'Composite_Score']])
