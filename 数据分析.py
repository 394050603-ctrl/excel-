import streamlit as st
import pandas as pd
import plotly.express as px

# 页面配置
st.set_page_config(
    page_title="订单/结算数据分析工具",
    page_icon="📊",
    layout="wide"
)

# 标题
st.title("📊 订单/结算数据分析工具")
st.markdown("---")

# 上传文件
uploaded_files = st.file_uploader("📁 上传你的Excel文件（支持多选）", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # 合并所有上传的Excel数据
    df_list = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        df_list.append(df)
    combined_df = pd.concat(df_list, ignore_index=True)
    
    st.success(f"✅ 成功上传并合并 {len(uploaded_files)} 个文件，共 {len(combined_df)} 条数据")
    
    # 数据预览
    st.subheader("🔍 数据预览（已清洗）")
    # 自动去重
    cleaned_df = combined_df.drop_duplicates()
    st.dataframe(cleaned_df, use_container_width=True)
    
    # 多维度分析
    st.markdown("---")
    st.subheader("📈 多维度数据分析")
    
    # 1. 按交易类型统计
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🧩 交易类型分布")
        type_counts = cleaned_df["交易类型"].value_counts().reset_index()
        type_counts.columns = ["交易类型", "订单数量"]
        fig_type = px.pie(type_counts, values="订单数量", names="交易类型", hole=0.3)
        st.plotly_chart(fig_type, use_container_width=True)
    
    # 2. 按时间趋势统计
    with col2:
        st.markdown("### 📅 订单量时间趋势")
        # 转换日期格式
        cleaned_df["下单时间"] = pd.to_datetime(cleaned_df["下单时间"])
        cleaned_df["月份"] = cleaned_df["下单时间"].dt.to_period("M").astype(str)
        time_trend = cleaned_df.groupby("月份")["订单号"].count().reset_index()
        time_trend.columns = ["月份", "订单数量"]
        fig_time = px.line(time_trend, x="月份", y="订单数量", markers=True)
        st.plotly_chart(fig_time, use_container_width=True)
    
    # 3. 按门店统计
    st.markdown("### 🏪 各门店数据对比")
    store_stats = cleaned_df.groupby("门店名称").agg({
        "订单号": "count",
        "商家应收（结算金额）": "sum"
    }).reset_index()
    store_stats.columns = ["门店名称", "订单数量", "应收总金额"]
    fig_store = px.bar(store_stats, x="门店名称", y="应收总金额", color="订单数量", barmode="group")
    st.plotly_chart(fig_store, use_container_width=True)
    
    # 数据导出
    st.markdown("---")
    st.subheader("💾 导出分析结果")
    col1, col2 = st.columns(2)
    with col1:
        csv = cleaned_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 下载清洗后的数据（CSV）",
            data=csv,
            file_name="清洗后订单数据.csv",
            mime="text/csv"
        )
    with col2:
        excel = cleaned_df.to_excel(index=False, engine="openpyxl")
        st.download_button(
            label="📥 下载清洗后的数据（Excel）",
            data=excel,
            file_name="清洗后订单数据.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("ℹ️ 请上传你的Excel文件开始分析，支持同时上传多个文件")