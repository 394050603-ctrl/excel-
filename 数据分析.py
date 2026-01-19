import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from datetime import datetime
import re
import numpy as np
from typing import Dict, List, Optional
import gc  # 垃圾回收
from streamlit.runtime.caching import cache_data

# ---------------------- 全局配置 & 性能优化基础 ----------------------
# 设置中文显示
pio.renderers.default = "browser"
px.defaults.template = "plotly_white"
px.defaults.color_continuous_scale = px.colors.sequential.Reds

# 页面配置
st.set_page_config(
    page_title="高性能多文件+自然语言数据分析工具",
    page_icon="📊",
    layout="wide"
)

# 性能参数配置
MAX_PREVIEW_ROWS = 100  # 预览最大行数
MAX_CHART_POINTS = 5000  # 图表最大数据点
CHUNK_SIZE = 10000       # 分块读取大小
CACHE_TTL = 3600         # 缓存有效期（秒）

# ---------------------- 缓存 & 高性能函数 ----------------------
@cache_data(ttl=CACHE_TTL)  # 缓存数据处理结果
def load_and_clean_data(uploaded_files: List) -> Optional[pd.DataFrame]:
    """
    高性能加载并清理多文件数据
    - 分块读取
    - 数据类型优化
    - 缺失值/重复值处理
    """
    df_list = []
    for file in uploaded_files:
        try:
            # 分块读取大文件
            if file.size > 10 * 1024 * 1024:  # 大于10MB的文件分块
                if file.name.endswith(".xlsx"):
                    chunks = pd.read_excel(file, chunksize=CHUNK_SIZE)
                else:
                    chunks = pd.read_csv(file, chunksize=CHUNK_SIZE)
                
                temp_df = pd.concat(chunks, ignore_index=True)
            else:  # 小文件直接读取
                if file.name.endswith(".xlsx"):
                    temp_df = pd.read_excel(file)
                else:
                    temp_df = pd.read_csv(file)
            
            # 添加来源文件列
            temp_df["来源文件"] = file.name
            
            # 数据类型优化（核心！减少内存占用）
            temp_df = optimize_dtypes(temp_df)
            
            # 预处理：缺失值/重复值
            temp_df = preprocess_data(temp_df)
            
            df_list.append(temp_df)
            
            # 释放内存
            del temp_df
            gc.collect()
            
        except Exception as e:
            st.warning(f"⚠️ 文件 {file.name} 读取失败：{str(e)}")
            continue
    
    if not df_list:
        return None
    
    # 合并数据并再次优化
    df = pd.concat(df_list, ignore_index=True)
    df = optimize_dtypes(df)
    
    # 释放临时列表内存
    del df_list
    gc.collect()
    
    return df

def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """优化数据类型，减少内存占用"""
    df_optimized = df.copy()
    
    # 数值列优化：int64→int32，float64→float32（无精度损失时）
    for col in df_optimized.select_dtypes(include=["int64"]).columns:
        if df_optimized[col].max() <= np.iinfo(np.int32).max and df_optimized[col].min() >= np.iinfo(np.int32).min:
            df_optimized[col] = df_optimized[col].astype(np.int32)
    
    for col in df_optimized.select_dtypes(include=["float64"]).columns:
        df_optimized[col] = df_optimized[col].astype(np.float32)
    
    # 字符串列优化：高频重复值→category
    for col in df_optimized.select_dtypes(include=["object"]).columns:
        if col != "来源文件" and len(df_optimized[col].unique()) / len(df_optimized[col]) < 0.1:  # 唯一值占比<10%
            df_optimized[col] = df_optimized[col].astype("category")
    
    # 日期列自动识别
    for col in df_optimized.columns:
        if col.lower() in ["日期", "时间", "date", "time"] and not pd.api.types.is_datetime64_any_dtype(df_optimized[col]):
            try:
                df_optimized[col] = pd.to_datetime(df_optimized[col], errors="coerce")
            except:
                pass
    
    return df_optimized

def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
    """预处理：缺失值/重复值处理"""
    # 删除全空行/列
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
    
    # 删除重复行
    df = df.drop_duplicates()
    
    # 数值列缺失值填充（用中位数，避免均值受异常值影响）
    for col in df.select_dtypes(include=["int32", "float32"]).columns:
        df[col] = df[col].fillna(df[col].median())
    
    return df

@cache_data(ttl=CACHE_TTL)
def parse_prompt_optimized(prompt: str, numeric_cols: List, category_cols: List, date_cols: List) -> Dict:
    """优化的指令解析函数（缓存解析结果）"""
    # 提取指标（数值列）
    target = None
    for col in numeric_cols:
        if col in prompt:
            target = col
            break
    
    # 提取分组维度
    group = None
    for col in category_cols + date_cols + ["来源文件"]:
        if col in prompt:
            group = col
            break
    
    # 提取图表类型
    chart_map = {
        "柱状图": "bar", "折线图": "line", "饼图": "pie",
        "散点图": "scatter", "箱线图": "box"
    }
    chart = None
    for key, val in chart_map.items():
        if key in prompt:
            chart = val
            break
    
    # 提取统计类型
    stat_map = {
        "总和": "sum", "平均值": "mean", "均值": "mean",
        "中位数": "median", "最大值": "max", "最小值": "min",
        "数量": "count", "计数": "count"
    }
    stat = "sum"
    for key, val in stat_map.items():
        if key in prompt:
            stat = val
            break
    
    # 提取筛选条件（向量化正则匹配）
    filters = []
    # 日期筛选
    for col in date_cols:
        if col in prompt:
            year_match = re.search(r"(\d{4})年", prompt)
            if year_match:
                filters.append((col, "year", int(year_match.group(1))))
            month_match = re.search(r"(\d{1,2})月", prompt)
            if month_match:
                filters.append((col, "month", int(month_match.group(1))))
            break
    
    # TopN/BottomN筛选
    top_match = re.search(r"最高(\d+)个", prompt)
    if top_match:
        filters.append(("top_n", int(top_match.group(1))))
    bottom_match = re.search(r"最低(\d+)个", prompt)
    if bottom_match:
        filters.append(("bottom_n", int(bottom_match.group(1))))
    
    return {
        "target": target,
        "group": group,
        "chart": chart,
        "stat": stat,
        "filters": filters
    }

def apply_filters_optimized(df: pd.DataFrame, filters: List, target_col: str) -> pd.DataFrame:
    """优化的筛选函数（先筛选后计算，向量化操作）"""
    filtered_df = df.copy()
    
    for filt in filters:
        if isinstance(filt, tuple):
            col, filt_type, val = filt
            if filt_type == "year":
                filtered_df = filtered_df[filtered_df[col].dt.year == val]
            elif filt_type == "month":
                filtered_df = filtered_df[filtered_df[col].dt.month == val]
        elif filt[0] == "top_n":
            n = filt[1]
            filtered_df = filtered_df.nlargest(n, target_col)
        elif filt[0] == "bottom_n":
            n = filt[1]
            filtered_df = filtered_df.nsmallest(n, target_col)
    
    # 限制最大数据量（避免图表卡顿）
    if len(filtered_df) > MAX_CHART_POINTS:
        filtered_df = filtered_df.sample(n=MAX_CHART_POINTS, random_state=42)
    
    return filtered_df

# ---------------------- 主界面逻辑 ----------------------
def main():
    st.title("📊 高性能多文件 + 自然语言指令数据分析应用")
    st.divider()
    
    # 第一步：批量上传数据
    st.subheader("1. 批量上传数据文件")
    uploaded_files = st.file_uploader(
        "支持同时上传多个Excel(.xlsx) / CSV(.csv)格式文件（支持10万行+大文件）",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        help="所有文件需包含相同表头，比如：日期、地区、产品、销售额、利润等"
    )
    
    if uploaded_files:
        with st.spinner("📥 正在加载并优化数据..."):
            df = load_and_clean_data(uploaded_files)
        
        if df is None:
            st.error("❌ 所有文件读取失败，请检查文件格式是否正确")
            return
        
        # 数据概览（轻量化预览）
        st.success(f"✅ 成功合并 {len(uploaded_files)} 个文件，共 {df.shape[0]:,} 行数据！")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总行数", f"{df.shape[0]:,}")
        with col2:
            st.metric("总列数", df.shape[1])
        with col3:
            st.metric("缺失值总数", df.isnull().sum().sum())
        with col4:
            mem_usage = df.memory_usage(deep=True).sum() / 1024 / 1024
            st.metric("内存占用", f"{mem_usage:.2f} MB")
        
        # 抽样预览（避免全量加载）
        st.subheader("2. 数据抽样预览（前100行）")
        preview_df = df.head(MAX_PREVIEW_ROWS)
        st.dataframe(preview_df, use_container_width=True)
        
        st.divider()
        
        # 第二步：自然语言指令输入
        st.subheader("3. 输入你的分析要求（自然语言）")
        st.info("💡 示例：计算各地区销售额总和并生成柱状图；找出2025年1月利润最高的3个产品；按月份统计销售额趋势并生成折线图")
        
        # 提取列名分类（基于优化后的数据）
        numeric_cols = df.select_dtypes(include=["int32", "float32", "int64", "float64"]).columns.tolist()
        category_cols = df.select_dtypes(include=["category", "object"]).columns.tolist()
        date_cols = df.select_dtypes(include=["datetime64"]).columns.tolist()
        
        # 指令输入
        user_prompt = st.text_area(
            "请输入你的分析需求",
            height=100,
            placeholder="例如：计算各地区销售额总和并生成柱状图；找出2025年1月利润最高的3个产品"
        )
        
        # 执行分析
        if st.button("📊 执行分析", type="primary") and user_prompt:
            with st.spinner("🔍 正在解析指令并分析数据..."):
                # 解析指令（缓存结果）
                parsed = parse_prompt_optimized(user_prompt, numeric_cols, category_cols, date_cols)
                target_col = parsed["target"]
                group_col = parsed["group"]
                chart_type = parsed["chart"]
                stat_type = parsed["stat"]
                filters = parsed["filters"]
                
                # 参数校验（提前终止无效计算）
                if not target_col:
                    st.error("❌ 未识别到分析指标，请确保输入中包含数值列名（如销售额、利润等）")
                    return
                if not group_col:
                    st.error("❌ 未识别到分组维度，请确保输入中包含分类/日期列名（如地区、日期等）")
                    return
                
                # 应用筛选（先筛选后计算，减少计算量）
                filtered_df = apply_filters_optimized(df, filters, target_col)
                
                # 批量聚合计算（一次groupby完成，避免多次计算）
                stats_df = filtered_df.groupby(group_col)[target_col].agg([
                    "sum", "mean", "median", "max", "min", "count"
                ]).round(2).reset_index()
                
                # 提取需要的统计值
                result_df = stats_df[[group_col, stat_type]].rename(columns={stat_type: target_col})
                
                # 生成轻量化图表
                fig = None
                if chart_type:
                    # 限制图表数据点
                    chart_df = result_df.head(MAX_CHART_POINTS)
                    
                    if chart_type == "bar":
                        fig = px.bar(chart_df, x=group_col, y=target_col, 
                                     title=f"{group_col} - {target_col} {stat_type} 柱状图", 
                                     text_auto=True, height=500)
                    elif chart_type == "line":
                        fig = px.line(chart_df, x=group_col, y=target_col, 
                                      title=f"{group_col} - {target_col} {stat_type} 趋势图", 
                                      markers=True, height=500)
                    elif chart_type == "pie":
                        fig = px.pie(chart_df, values=target_col, names=group_col, 
                                     title=f"{group_col} - {target_col} {stat_type} 占比图", 
                                     hole=0.3, height=500)
                    elif chart_type == "scatter":
                        fig = px.scatter(filtered_df.head(MAX_CHART_POINTS), x=group_col, y=target_col, 
                                         title=f"{group_col} - {target_col} 散点图", 
                                         color=group_col, height=500)
                    elif chart_type == "box":
                        fig = px.box(filtered_df.head(MAX_CHART_POINTS), x=group_col, y=target_col, 
                                     title=f"{group_col} - {target_col} 箱线图", height=500)
                
                # 展示结果
                st.divider()
                st.subheader("4. 分析结果")
                
                # 表格结果
                st.dataframe(result_df, use_container_width=True)
                
                # 图表结果
                if fig:
                    fig.update_layout(xaxis_title=group_col, yaxis_title=target_col)
                    st.plotly_chart(fig, use_container_width=True)
                
                # 导出结果（轻量化）
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_filename = f"分析结果_{timestamp}.xlsx"
                st.download_button(
                    label="📥 下载分析结果（Excel）",
                    data=result_df.to_excel(index=False, engine="openpyxl"),
                    file_name=export_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 释放临时数据内存
                del filtered_df, stats_df, result_df
                gc.collect()
    
    else:
        # 未上传文件提示
        st.info("💡 请批量上传Excel/CSV数据文件（支持10万行+大文件），即可开始高性能数据分析")
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

if __name__ == "__main__":
    main()
