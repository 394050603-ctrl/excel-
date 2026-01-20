import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="智能表格数据分析工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 智能自适应表格数据分析工具")
st.markdown("### 支持任意格式表格自动解析，无行列/格式限制")

# ---------------------- 核心：智能表格解析函数 ----------------------
def smart_parse_excel(file, sheet_name=None):
    """
    智能解析Excel文件，自动定位有效数据区域，兼容任意格式
    """
    # 读取所有sheet（如果未指定）
    if sheet_name is None:
        xl_file = pd.ExcelFile(file)
        sheet_names = xl_file.sheet_names
        all_data = {}
        for name in sheet_names:
            df = parse_single_sheet(file, name)
            if not df.empty:
                all_data[name] = df
        return all_data
    else:
        df = parse_single_sheet(file, sheet_name)
        return {sheet_name: df}

def parse_single_sheet(file, sheet_name):
    """
    解析单个sheet，自动定位有效数据、处理合并单元格、空行空列
    """
    # 先用openpyxl读取原始表格（处理合并单元格）
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[sheet_name]
    
    # 步骤1：定位有效数据区域（跳过全空的行/列）
    min_row, max_row = ws.min_row, ws.max_row
    min_col, max_col = ws.min_col, ws.max_col
    
    # 过滤全空行
    valid_rows = []
    for row in range(min_row, max_row + 1):
        row_data = [ws.cell(row=row, column=col).value for col in range(min_col, max_col + 1)]
        if any(cell is not None and str(cell).strip() != "" for cell in row_data):
            valid_rows.append(row)
    
    # 过滤全空列
    valid_cols = []
    for col in range(min_col, max_col + 1):
        col_data = [ws.cell(row=row, column=col).value for row in valid_rows]
        if any(cell is not None and str(cell).strip() != "" for cell in col_data):
            valid_cols.append(col)
    
    if not valid_rows or not valid_cols:
        return pd.DataFrame()
    
    # 步骤2：提取有效数据（处理合并单元格的值填充）
    data = []
    header_row = valid_rows[0]  # 第一行作为表头（自动适配）
    data_rows = valid_rows[1:]  # 其余行作为数据
    
    # 提取表头
    headers = []
    for col in valid_cols:
        cell = ws.cell(row=header_row, column=col)
        header = cell.value if cell.value is not None else f"列{get_column_letter(col)}"
        headers.append(str(header).strip())
    
    # 提取数据行（填充合并单元格的值）
    for row in data_rows:
        row_vals = []
        for col in valid_cols:
            cell = ws.cell(row=row, column=col)
            # 如果单元格是合并的，取合并区域的第一个值
            if cell.coordinate in ws.merged_cells:
                for merged_range in ws.merged_cells:
                    if cell.coordinate in merged_range:
                        merged_cell = ws[merged_range.split(":")[0]]
                        row_vals.append(merged_cell.value)
                        break
            else:
                row_vals.append(cell.value)
        data.append(row_vals)
    
    # 步骤3：构建DataFrame并清洗
    df = pd.DataFrame(data, columns=headers)
    # 清洗空值、空白字符串
    df = df.replace("", None).dropna(how="all")
    # 自动转换数值列（文本型数字转数值）
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col], errors="ignore")
        except:
            pass
    
    return df

# ---------------------- 上传文件 & 解析 ----------------------
uploaded_files = st.file_uploader(
    "上传Excel文件（支持多个）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    for file in uploaded_files:
        st.markdown(f"### 📄 解析文件：{file.name}")
        
        # 智能解析表格
        with st.spinner("正在智能解析表格数据..."):
            sheet_data_dict = smart_parse_excel(file)
        
        if not sheet_data_dict:
            st.warning(f"文件{file.name}未识别到有效数据，请检查表格内容")
            continue
        
        # 选择要分析的sheet（自动列出所有有数据的sheet）
        sheet_names = list(sheet_data_dict.keys())
        selected_sheet = st.selectbox(
            f"选择{file.name}的sheet",
            sheet_names,
            key=f"sheet_{file.name}"
        )
        
        # 获取解析后的DataFrame
        df = sheet_data_dict[selected_sheet]
        st.markdown(f"#### 📋 解析结果预览（自动适配{selected_sheet}）")
        st.dataframe(df, use_container_width=True)
        
        # ---------------------- 自动分析 ----------------------
        st.markdown(f"#### 📈 自动数据分析（无格式/行列限制）")
        
        # 自动识别数值列
        numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns.tolist()
        if not numeric_cols:
            st.info("未识别到数值列，仅展示文本列统计")
            # 文本列统计
            text_cols = df.select_dtypes(include=["object"]).columns.tolist()
            if text_cols:
                selected_text_col = st.selectbox("选择文本列分析", text_cols, key=f"text_{file.name}")
                # 文本列频次统计
                text_counts = df[selected_text_col].value_counts().reset_index()
                text_counts.columns = [selected_text_col, "频次"]
                
                # 可视化
                fig = px.bar(
                    text_counts.head(20),  # 取前20个高频值
                    x=selected_text_col,
                    y="频次",
                    title=f"{selected_text_col} - 频次分布",
                    template="plotly_white"
                )
                st.plotly_chart(fig, use_container_width=True)
        else:
            # 数值列分析
            selected_num_col = st.selectbox(
                "选择数值列分析",
                numeric_cols,
                key=f"num_{file.name}"
            )
            
            # 基础统计信息
            stats = df[selected_num_col].describe()
            st.markdown("##### 📊 基础统计")
            st.dataframe(stats, use_container_width=True)
            
            # 可视化（自动适配）
            col1, col2 = st.columns(2)
            with col1:
                # 直方图
                fig_hist = px.histogram(
                    df,
                    x=selected_num_col,
                    title=f"{selected_num_col} - 分布直方图",
                    template="plotly_white"
                )
                st.plotly_chart(fig_hist, use_container_width=True)
            
            with col2:
                # 箱线图（看异常值）
                fig_box = px.box(
                    df,
                    y=selected_num_col,
                    title=f"{selected_num_col} - 箱线图（异常值分析）",
                    template="plotly_white"
                )
                st.plotly_chart(fig_box, use_container_width=True)
            
            # 可选：按文本列分组分析
            text_cols = df.select_dtypes(include=["object"]).columns.tolist()
            if text_cols:
                selected_group_col = st.selectbox(
                    "选择分组列（可选）",
                    ["无"] + text_cols,
                    key=f"group_{file.name}"
                )
                if selected_group_col != "无":
                    # 分组统计
                    group_stats = df.groupby(selected_group_col)[selected_num_col].agg(["mean", "sum", "count"]).reset_index()
                    st.markdown(f"##### 📈 按{selected_group_col}分组统计")
                    st.dataframe(group_stats, use_container_width=True)
                    
                    # 分组可视化
                    fig_group = px.bar(
                        group_stats,
                        x=selected_group_col,
                        y="sum",
                        title=f"{selected_group_col} - {selected_num_col} 总和",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_group, use_container_width=True)

else:
    st.info("✅ 请上传你的Excel文件（任意格式），工具会自动检索表格数据并分析")
