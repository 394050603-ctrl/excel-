import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from datetime import datetime

# 设置中文显示
pio.renderers.default = "browser"
px.defaults.template = "plotly_white"
px.defaults.color_continuous_scale = px.colors.sequential.Reds

# 页面配置
st.set_page_config(
    page_title="自定义数据分析工具",
    page_icon="📊",
    layout="wide"
)

# 标题
st.title("📊 自定义条件数据分析应用")
st.divider()

# ---------------------- 第一步：上传数据 ----------------------
st.subheader("1. 上传数据文件")
uploaded_file = st.file_uploader(
    "支持Excel(.xlsx) / CSV(.csv)格式",
    type=["xlsx", "csv"],
    help="请确保文件有表头，比如：日期、地区、产品、销售额、利润等"
)

if uploaded_file is not None:
    # 读取数据
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)
        
        # 数据预览
        st.success("✅ 数据读取成功！")
        st.subheader("数据预览")
        st.dataframe(df.head(10), use_container_width=True)
        
        # 显示数据基本信息
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("总行数", df.shape[0])
        with col2:
            st.metric("总列数", df.shape[1])
        with col3:
            st.metric("缺失值总数", df.isnull().sum().sum())
        
        st.divider()

        # ---------------------- 第二步：选择分析维度 ----------------------
        st.subheader("2. 选择分析维度")
        
        # 提取列名并分类
        numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns.tolist()  # 数值列（销售额、利润等）
        category_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()  # 分类列（地区、产品等）
        date_cols = df.select_dtypes(include=["datetime64"]).columns.tolist()  # 日期列
        
        # 自动识别日期列（如果是字符串格式）
        for col in df.columns:
            if col.lower() in ["日期", "时间", "date", "time"] and col not in date_cols:
                try:
                    df[col] = pd.to_datetime(df[col])
                    date_cols.append(col)
                except:
                    pass

        # 选择分析指标（Y轴：数值）
        target_col = st.selectbox(
            "📌 要分析的核心指标（比如销售额、利润）",
            numeric_cols,
            index=0 if numeric_cols else None,
            help="选择需要计算/可视化的数值字段"
        )

        # 选择分组维度（X轴：分类/日期）
        group_col = st.selectbox(
            "📌 分组维度（比如地区、产品、日期）",
            category_cols + date_cols,
            index=0 if (category_cols + date_cols) else None,
            help="按哪个维度对指标进行分组分析"
        )

        # ---------------------- 第三步：设置筛选条件 ----------------------
        st.subheader("3. 设置筛选条件（可选）")
        filter_options = st.expander("🔍 展开设置筛选条件", expanded=False)
        
        with filter_options:
            # 多条件筛选
            filters = []
            # 1. 数值筛选
            if numeric_cols:
                filter_num_col = st.selectbox("选择数值筛选列", numeric_cols)
                filter_num_oper = st.selectbox("筛选条件", ["大于", "小于", "等于", "大于等于", "小于等于"])
                filter_num_val = st.number_input(f"值（{filter_num_col}）", value=0)
                
                oper_map = {
                    "大于": ">", "小于": "<", "等于": "==",
                    "大于等于": ">=", "小于等于": "<="
                }
                filters.append(f"df['{filter_num_col}'] {oper_map[filter_num_oper]} {filter_num_val}")
            
            # 2. 分类筛选
            if category_cols:
                filter_cat_col = st.selectbox("选择分类筛选列", category_cols)
                filter_cat_vals = st.multiselect(
                    f"选择{filter_cat_col}的取值",
                    df[filter_cat_col].unique()
                )
                if filter_cat_vals:
                    filters.append(f"df['{filter_cat_col}'].isin({filter_cat_vals})")
            
            # 3. 日期筛选
            if date_cols:
                filter_date_col = st.selectbox("选择日期筛选列", date_cols)
                date_start = st.date_input("开始日期", value=df[filter_date_col].min())
                date_end = st.date_input("结束日期", value=df[filter_date_col].max())
                filters.append(f"df['{filter_date_col}'] >= '{date_start}'")
                filters.append(f"df['{filter_date_col}'] <= '{date_end}'")
        
        # 应用筛选条件
        filtered_df = df.copy()
        if filters:
            filter_expr = " & ".join(filters)
            try:
                filtered_df = filtered_df.query(filter_expr)
                st.info(f"🔎 筛选后剩余数据：{len(filtered_df)} 行")
            except:
                st.warning("⚠️ 筛选条件设置有误，将使用原始数据")
        
        # ---------------------- 第四步：选择分析类型 ----------------------
        st.subheader("4. 选择分析类型")
        analysis_type = st.radio(
            "",
            [
                "📊 基础统计分析（求和/均值/中位数）",
                "📈 可视化分析（图表）",
                "🎯 异常值检测",
                "📄 导出分析结果"
            ],
            horizontal=True
        )

        # ---------------------- 执行分析 ----------------------
        st.divider()
        st.subheader("5. 分析结果")

        # 1. 基础统计分析
        if "基础统计分析" in analysis_type:
            stats_df = filtered_df.groupby(group_col)[target_col].agg([
                "count", "sum", "mean", "median", "max", "min", "std"
            ]).round(2)
            stats_df.columns = ["数量", "总和", "平均值", "中位数", "最大值", "最小值", "标准差"]
            
            st.dataframe(stats_df, use_container_width=True)
            
            # 关键指标高亮
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(f"{target_col}总和", f"{stats_df['总和'].sum():,.2f}")
            with col2:
                st.metric(f"{target_col}平均值", f"{stats_df['平均值'].mean():,.2f}")
            with col3:
                st.metric(f"最大{target_col}分组", stats_df['最大值'].idxmax())

        # 2. 可视化分析
        elif "可视化分析" in analysis_type:
            chart_type = st.selectbox(
                "选择图表类型",
                ["柱状图", "折线图", "饼图", "散点图", "箱线图"]
            )
            
            # 生成图表
            fig = None
            if chart_type == "柱状图":
                fig = px.bar(
                    filtered_df,
                    x=group_col,
                    y=target_col,
                    title=f"{group_col} - {target_col} 柱状图",
                    color=group_col,
                    text_auto=True
                )
            elif chart_type == "折线图":
                fig = px.line(
                    filtered_df,
                    x=group_col,
                    y=target_col,
                    title=f"{group_col} - {target_col} 趋势图",
                    markers=True
                )
            elif chart_type == "饼图":
                pie_data = filtered_df.groupby(group_col)[target_col].sum().reset_index()
                fig = px.pie(
                    pie_data,
                    values=target_col,
                    names=group_col,
                    title=f"{group_col} - {target_col} 占比图",
                    hole=0.3
                )
            elif chart_type == "散点图":
                fig = px.scatter(
                    filtered_df,
                    x=group_col,
                    y=target_col,
                    title=f"{group_col} - {target_col} 散点图",
                    color=group_col,
                    size=target_col
                )
            elif chart_type == "箱线图":
                fig = px.box(
                    filtered_df,
                    x=group_col,
                    y=target_col,
                    title=f"{group_col} - {target_col} 箱线图（异常值分析）"
                )
            
            if fig:
                fig.update_layout(height=600, xaxis_title=group_col, yaxis_title=target_col)
                st.plotly_chart(fig, use_container_width=True)

        # 3. 异常值检测
        elif "异常值检测" in analysis_type:
            st.info(f"📌 基于{target_col}的异常值检测（四分位法）")
            
            # 计算四分位数
            q1 = filtered_df[target_col].quantile(0.25)
            q3 = filtered_df[target_col].quantile(0.75)
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            
            # 识别异常值
            outliers = filtered_df[
                (filtered_df[target_col] < lower_bound) | (filtered_df[target_col] > upper_bound)
            ]
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("异常值数量", len(outliers))
                st.metric("正常范围", f"{lower_bound:.2f} ~ {upper_bound:.2f}")
            with col2:
                st.metric("最小值", filtered_df[target_col].min())
                st.metric("最大值", filtered_df[target_col].max())
            
            # 展示异常值
            if len(outliers) > 0:
                st.subheader("异常值详情")
                st.dataframe(outliers, use_container_width=True)
                
                # 异常值可视化
                fig = px.box(
                    filtered_df,
                    y=target_col,
                    title=f"{target_col} 异常值分布",
                    color_discrete_sequence=["#FF4B4B"]
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("✅ 未检测到异常值！")

        # 4. 导出分析结果
        elif "导出分析结果" in analysis_type:
            # 生成导出数据
            export_df = filtered_df.groupby(group_col)[target_col].agg([
                "sum", "mean", "median", "max", "min"
            ]).round(2).reset_index()
            
            # 导出为Excel
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_filename = f"数据分析结果_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 下载分析结果（Excel）",
                data=export_df.to_excel(index=False, engine="openpyxl"),
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 预览导出数据
            st.subheader("导出数据预览")
            st.dataframe(export_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ 数据处理出错：{str(e)}")
        st.info("请检查文件格式是否正确，或表头是否规范")

else:
    # 未上传文件时的提示
    st.info("💡 请先上传Excel/CSV数据文件，即可开始自定义数据分析")
    # 示例数据预览
    with st.expander("查看示例数据格式"):
        sample_data = pd.DataFrame({
            "日期": pd.date_range(start="2025-01-01", periods=10),
            "地区": ["华东", "华北", "华南"] * 3 + ["华东"],
            "产品类别": ["电子产品", "日用品", "食品"] * 3 + ["电子产品"],
            "销售额": [12000, 8000, 5000, 15000, 9000, 6000, 13000, 7000, 4500, 14000],
            "利润": [2400, 1600, 1000, 3000, 1800, 1200, 2600, 1400, 900, 2800]
        })
        st.dataframe(sample_data, use_container_width=True)
