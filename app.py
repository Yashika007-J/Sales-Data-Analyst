import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data
df = pd.read_csv(r"D:\SALES_PROJECT\Sales-Data-Analyst\coffee_data.csv")

# Title
st.title("☕ Afficionado Coffee Roasters - Sales Analytics")
st.markdown("**Time-Based Performance Analysis Dashboard**")

# Sidebar filters
st.sidebar.header("Filters")
store = st.sidebar.multiselect(
    "Select Store Location",
    options=df['store_location'].unique(),
    default=df['store_location'].unique()
)
hour_range = st.sidebar.slider("Select Hour Range", 6, 20, (6, 20))

# Filter data
filtered_df = df[
    (df['store_location'].isin(store)) &
    (df['hour'] >= hour_range[0]) &
    (df['hour'] <= hour_range[1])
]

# KPI Cards
col1, col2, col3 = st.columns(3)
col1.metric("Total Revenue", f"${filtered_df['revenue'].sum():,.2f}")
col2.metric("Total Transactions", f"{len(filtered_df):,}")
col3.metric("Avg Revenue/Transaction", f"${filtered_df['revenue'].mean():,.2f}")

st.divider()

# Chart 1 - Hourly transactions
st.subheader("📈 Transaction Volume by Hour")
hour_count = filtered_df.groupby('hour')['transaction_id'].count()
fig, ax = plt.subplots(figsize=(10, 4))
ax.plot(hour_count.index, hour_count.values, color='#6F4E37', linewidth=2, marker='o')
ax.fill_between(hour_count.index, hour_count.values, alpha=0.3, color='#6F4E37')
ax.set_xlabel("Hour of Day")
ax.set_ylabel("Transactions")
st.pyplot(fig)

st.divider()

# Chart 2 - Revenue by store
st.subheader("🏪 Revenue by Store Location")
store_rev = filtered_df.groupby('store_location')['revenue'].sum()
fig2, ax2 = plt.subplots(figsize=(8, 4))
ax2.bar(store_rev.index, store_rev.values, color=['#6F4E37','#D2691E','#A0522D'])
ax2.set_xlabel("Store Location")
ax2.set_ylabel("Total Revenue ($)")
st.pyplot(fig2)

st.divider()

# Chart 3 - Time bucket
st.subheader("⏰ Revenue by Time of Day")
bucket_rev = filtered_df.groupby('time_bucket')['revenue'].sum()
fig3, ax3 = plt.subplots(figsize=(8, 4))
ax3.bar(bucket_rev.index, bucket_rev.values, color='#8B4513')
ax3.set_xlabel("Time Bucket")
ax3.set_ylabel("Total Revenue ($)")
plt.xticks(rotation=15)
st.pyplot(fig3)