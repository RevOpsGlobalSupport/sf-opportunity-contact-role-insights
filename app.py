import altair as alt

st.subheader("Visual Insights")

# Create a lightweight stage group for charts
chart_df = opps.copy()
chart_df["Stage Group"] = "Open"
chart_df.loc[won_mask, "Stage Group"] = "Won"
chart_df.loc[lost_mask, "Stage Group"] = "Lost"

# -----------------------
# 1) Avg Contact Roles by Outcome
# -----------------------
avg_contacts_df = (
    chart_df.groupby("Stage Group")["contact_count"]
    .mean()
    .reset_index()
    .rename(columns={"contact_count": "Avg Contact Roles"})
)

chart1 = (
    alt.Chart(avg_contacts_df)
    .mark_bar()
    .encode(
        x=alt.X("Stage Group:N", sort=["Won", "Open", "Lost"], title="Opportunity Outcome"),
        y=alt.Y("Avg Contact Roles:Q", title="Avg Contact Roles"),
        tooltip=["Stage Group", alt.Tooltip("Avg Contact Roles:Q", format=".2f")]
    )
    .properties(height=280)
)
st.altair_chart(chart1, use_container_width=True)

# -----------------------
# 2) Pipeline by Contact Coverage Buckets
# -----------------------
bucket_df = chart_df.copy()
bucket_df["Coverage Bucket"] = pd.cut(
    bucket_df["contact_count"],
    bins=[-1, 0, 1, 1000],
    labels=["0 Contact Roles", "1 Contact Role", "2+ Contact Roles"]
)

pipeline_bucket = (
    bucket_df.groupby(["Stage Group", "Coverage Bucket"])["Amount"]
    .sum()
    .reset_index()
)

chart2 = (
    alt.Chart(pipeline_bucket)
    .mark_bar()
    .encode(
        x=alt.X("Coverage Bucket:N", title="Contact Role Coverage"),
        y=alt.Y("Amount:Q", title="Pipeline ($)"),
        color=alt.Color("Stage Group:N", title="Outcome"),
        tooltip=[
            "Stage Group", "Coverage Bucket",
            alt.Tooltip("Amount:Q", format=",.0f")
        ]
    )
    .properties(height=320)
)
st.altair_chart(chart2, use_container_width=True)

# -----------------------
# 3) Distribution of Contact Roles
# -----------------------
chart3 = (
    alt.Chart(chart_df)
    .mark_bar(opacity=0.8)
    .encode(
        x=alt.X("contact_count:Q", bin=alt.Bin(maxbins=15), title="Contact Roles per Opportunity"),
        y=alt.Y("count():Q", title="Number of Opportunities"),
        color=alt.Color("Stage Group:N", title="Outcome"),
        tooltip=[
            alt.Tooltip("contact_count:Q", title="Contact Roles"),
            alt.Tooltip("count():Q", title="Opp Count")
        ]
    )
    .properties(height=320)
)
st.altair_chart(chart3, use_container_width=True)

# -----------------------
# 4) Contact Roles vs Amount (Scatter)
# -----------------------
scatter_df = chart_df.copy()
scatter_df["Amount"] = pd.to_numeric(scatter_df["Amount"], errors="coerce").fillna(0)

chart4 = (
    alt.Chart(scatter_df)
    .mark_circle(size=70, opacity=0.55)
    .encode(
        x=alt.X("contact_count:Q", title="Contact Roles"),
        y=alt.Y("Amount:Q", title="Amount ($)"),
        color=alt.Color("Stage Group:N", title="Outcome"),
        tooltip=[
            "Opportunity ID",
            "Stage Group",
            alt.Tooltip("contact_count:Q", title="Contact Roles"),
            alt.Tooltip("Amount:Q", format=",.0f", title="Amount")
        ]
    )
    .properties(height=360)
)
st.altair_chart(chart4, use_container_width=True)
