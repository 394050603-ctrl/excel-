import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from datetime import datetime
import re

# 设置中文显示
pio.renderers.default = "browser"
px.defaults.template = "plotly_white"
px.defaults.color_continuous_scale = px.colors.sequential.Reds

# 页面配置
st.set_page_config(
    page_title="多文件 + 自然语言指令数据分析工具",
    page_icon="📊",
    layout="wide"
)

# 标题
st.title("📊 多文件上传 + 自然语言指令数据分析应用")
st.divider()

# ---------------------- 第一步：批量上传数据 ----------------------
st.subheader("1. 批量上传数据文件")
uploaded_files = st.file_uploader(
    "支持同时上传多个Excel(.xlsx) / CSV(.csv)格式文件",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="所有文件需包含相同表头，比如：日期、地区、产品、销售额、利润等"
)

if uploaded_files:
    # 合并所有上传的文件
    df_list = []
    for file in uploaded_files:
        try:
            if file.name.endswith(".xlsx"):
                temp_df = pd.read_excel(file)
            else:
                temp_df = pd.read_csv(file)
            # 添加来源文件列，方便追溯
            temp_df["来源文件"] = file.name
            df_list.append(temp_df)
        except Exception as e:
            st.warning(f"⚠️ 文件 {file.name} 读取失败：{str(e)}")
    
    if df_list:
        # 合并所有数据
        df = pd.concat(df_list, ignore_index=True)
        st.success(f"✅ 成功合并 {len(df_list)} 个文件，共 {df.shape[0]} 行数据！")
        
        # 数据预览
        st.subheader("合并后数据预览")
        st.dataframe(df.head(10), use_container_width=True)
        
        # 显示数据基本信息
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总行数", df.shape[0])
        with col2:
            st.metric("总列数", df.shape[1])
        with col3:
            st.metric("缺失值总数", df.isnull().sum().sum())
        with col4:
            st.metric("上传文件数", len(df_list))
        
        st.divider()

        # ---------------------- 第二步：自然语言输入分析要求 ----------------------
        st.subheader("2. 输入你的分析要求（自然语言）")
        st.info("💡 示例：计算各地区销售额总和并生成柱状图；找出2025年1月利润最高的3个产品；按月份统计销售额趋势并生成折线图")
        
        # 提取列名并分类
        numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns.tolist()
        category_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
        date_cols = df.select_dtypes(include=["datetime64"]).columns.tolist()
        
        # 自动识别日期列
        for col in df.columns:
            if col.lower() in ["日期", "时间", "date", "time"] and col not in date_cols:
                try:
                    df[col] = pd.to_datetime(df[col])
                    date_cols.append(col)
                except:
                    pass

        # 输入分析指令
        user_prompt = st.text_area(
            "请输入你的分析需求",
            height=100,
            placeholder="例如：计算各地区销售额总和并生成柱状图；找出2025年1月利润最高的3个产品"
        )

        # 解析用户指令
        def parse_prompt(prompt):
            # 提取指标（数值列）
            target = None
            for col in numeric_cols:
                if col in prompt:
                    target = col
                    break
            
            # 提取分组维度（分类/日期列）
            group = None
            for col in category_cols + date_cols + ["来源文件"]:
                if col in prompt:
                    group = col
                    break
            
            # 提取图表类型
            chart = None
            if "柱状图" in prompt:
                chart = "bar"
            elif "折线图" in prompt:
                chart = "line"
            elif "饼图" in prompt:
                chart = "pie"
            elif "散点图" in prompt:
                chart = "scatter"
            elif "箱线图" in prompt:
                chart = "box"
            
            # 提取统计类型
            stat = "sum"
            if "平均值" in prompt or "均值" in prompt:
                stat = "mean"
            elif "中位数" in prompt:
                stat = "median"
            elif "最大值" in prompt:
                stat = "max"
            elif "最小值" in prompt:
                stat = "min"
            elif "数量" in prompt or "计数" in prompt:
                stat = "count"
            
            # 提取筛选条件
            filters = []
            # 日期筛选
            for col in date_cols:
                if col in prompt:
                    # 匹配年份
                    year_match = re.search(r"(\d{4})年", prompt)
                    if year_match:
                        year = year_match.group(1)
                        filters.append(f"df['{col}'].dt.year == {year}")
                    # 匹配月份
                    month_match = re.search(r"(\d{1,2})月", prompt)
                    if month_match:
                        month = month_match.group(1)
                        filters.append(f"df['{col}'].dt.month == {month}")
                    break
            
            # 数值筛选
            if "最高" in prompt:
                top_n = re.search(r"最高(\d+)个", prompt)
                if top_n:
                    filters.append(f"top_{top_n.group(1)}")
            elif "最低" in prompt:
                bottom_n = re.search(r"最低(\d+)个", prompt)
                if bottom_n:
                    filters.append(f"bottom_{bottom_n.group(1)}")
            
            return {
                "target": target,
                "group": group,
                "chart": chart,
                "stat": stat,
                "filters": filters
            }

        # 执行分析
        if st.button("📊 执行分析") and user_prompt:
            with st.spinner("正在分析数据..."):
                parsed = parse_prompt(user_prompt)
                target_col = parsed["target"]
                group_col = parsed["group"]
                chart_type = parsed["chart"]
                stat_type = parsed["stat"]
                filters = parsed["filters"]

                if not target_col:
                    st.error("❌ 未识别到分析指标，请确保输入中包含数值列名（如销售额、利润等）")
                elif not group_col:
                    st.error("❌ 未识别到分组维度，请确保输入中包含分类/日期列名（如地区、日期等）")
                else:
                    # 应用筛选条件
                    filtered_df = df.copy()
                    for filt in filters:
                        if filt.startswith("top_"):
                            n = int(filt.split("_")[1])
                            filtered_df = filtered_df.nlargest(n, target_col)
                        elif filt.startswith("bottom_"):
                            n = int(filt.split("_")[1])
                            filtered_df = filtered_df.nsmallest(n, target_col)
                        else:
                            try:
                                filtered_df = filtered_df.query(filt)
                            except:
                                pass
                    
                    # 计算统计值
                    stats_df = filtered_df.groupby(group_col)[target_col].agg(stat_type).round(2).reset_index()
                    stats_df.columns = [group_col, target_col]

                    # 生成图表
                    fig = None
                    if chart_type == "bar":
                        fig = px.bar(stats_df, x=group_col, y=target_col, title=f"{group_col} - {target_col} {stat_type} 柱状图", text_auto=True)
                    elif chart_type == "line":
                        fig = px.line(stats_df, x=group_col, y=target_col, title=f"{group_col} - {target_col} {stat_type} 趋势图", markers=True)
                    elif chart_type == "pie":
                        fig = px.pie(stats_df, values=target_col, names=group_col, title=f"{group_col} - {target_col} {stat_type} 占比图", hole=0.3)
                    elif chart_type == "scatter":
                        fig = px.scatter(filtered_df, x=group_col, y=target_col, title=f"{group_col} - {target_col} 散点图", color=group_col)
                    elif chart_type == "box":
                        fig = px.box(filtered_df, x=group_col, y=target_col, title=f"{group_col} - {target_col} 箱线图")
                    
                    # 展示结果
                    st.divider()
                    st.subheader("3. 分析结果")
                    
                    # 表格结果
                    st.dataframe(stats_df, use_container_width=True)
                    
                    # 图表结果
                    if fig:
                        fig.update_layout(height=600, xaxis_title=group_col, yaxis_title=target_col)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    # 导出结果
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_filename = f"分析结果_{timestamp}.xlsx"
                    st.download_button(
                        label="📥 下载分析结果（Excel）",
                        data=stats_df.to_excel(index=False, engine="openpyxl"),
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    else:
        st.error("❌ 所有文件读取失败，请检查文件格式是否正确")

else:
    # 未上传文件时的提示
    st.info("💡 请批量上传Excel/CSV数据文件（支持多文件），即可开始自然语言指令数据分析")
    # 示例数据预览
    with st.expander("查看示例数据格式"):
        sample_data = pd.DataFrame({
            "日期": pd.date_range(start="2025-01-01", periods=10),
            "地区": ["华东", "华北", "华南"] * 3 + ["华东"],
            "产品类别": ["电子产品", "日用品", "食品"] * 3 + ["电子产品"],
            "销售额": [12000, 8000, 5000, 15000, 9000, 6000, 13000, 7000, 4500, 14000],
            "利润": [2400, 1600, 1000, 3000, 1800, 1200, 2600, 1400, 900, 2800],
            "来源文件": ["文件1.xlsx"] * 5 + ["文件2.xlsx"] * 5
        })
        st.dataframe(sample_data, use_container_width=True)
