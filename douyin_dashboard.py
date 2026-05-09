"""
抖音视频数据可视化看板（Streamlit Cloud 部署版）
数据文件：同目录下的 抖音视频数据汇总.xlsx
公网访问：https://your-app-name.streamlit.app
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import warnings
from pathlib import Path

warnings.filterwarnings('ignore')

# ======================== 页面配置 ========================
st.set_page_config(
    page_title="抖音视频数据分析看板",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================== 数据文件路径（相对路径） ========================
DATA_FILE = Path(__file__).parent / "抖音视频数据汇总.xlsx"

# ======================== 数据加载与清洗 ========================
@st.cache_data
def load_data(file_path):
    """从同目录加载Excel"""
    if not file_path.exists():
        st.error(f"❌ 文件不存在: {file_path}\n请确保已将数据文件上传到仓库根目录。")
        st.stop()
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        st.success(f"✅ 成功加载数据，共 {df.shape[0]} 行，{df.shape[1]} 列")
        return df
    except Exception as e:
        st.error(f"读取Excel失败: {e}")
        st.stop()

def standardize_columns(df):
    """将中文列名映射到标准英文字段"""
    col_mapping = {
        'author_name': '作者昵称',
        'author_id': '抖音号',
        'video_id': '视频ID',
        'description': '视频描述',
        'publish_date': '创建时间',
        'uid': '作者UID',
        'comment_count': '评论数',
        'like_count': '点赞数',
        'share_count': '分享数',
        'collect_count': '收藏数',
        'video_link': '视频链接',
        'fans_count': '粉丝数',
        'total_likes_received': '获赞总数',
        'sec_uid': '用户sec_user_id'
    }
    used_mapping = {orig: std for std, orig in col_mapping.items() if orig in df.columns}
    df_renamed = df.rename(columns=used_mapping)

    # 数值列转换
    num_cols = ['comment_count', 'like_count', 'share_count', 'collect_count', 'fans_count', 'total_likes_received']
    for col in num_cols:
        if col in df_renamed.columns:
            df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce')

    # 日期列转换
    if 'publish_date' in df_renamed.columns:
        df_renamed['publish_date'] = pd.to_datetime(df_renamed['publish_date'], errors='coerce')

    return df_renamed

def compute_kpis(df):
    """计算关键指标"""
    total_videos = len(df)
    total_likes = df['like_count'].sum() if 'like_count' in df else None
    total_comments = df['comment_count'].sum() if 'comment_count' in df else None
    total_shares = df['share_count'].sum() if 'share_count' in df else None
    total_collects = df['collect_count'].sum() if 'collect_count' in df else None
    avg_likes = df['like_count'].mean() if 'like_count' in df else None
    if 'author_name' in df and 'fans_count' in df:
        total_fans = df.groupby('author_name')['fans_count'].max().sum()
    else:
        total_fans = None
    return {
        'total_videos': total_videos,
        'total_likes': total_likes,
        'total_comments': total_comments,
        'total_shares': total_shares,
        'total_collects': total_collects,
        'avg_likes_per_video': avg_likes,
        'total_fans': total_fans
    }

# ======================== 图表绘制函数 ========================
def plot_trend(df, date_col, value_col, title, color='#FF5722'):
    if date_col not in df or value_col not in df:
        return None
    temp = df[[date_col, value_col]].dropna()
    if temp.empty:
        return None
    temp['date'] = temp[date_col].dt.date
    daily = temp.groupby('date')[value_col].sum().reset_index()
    fig = px.line(daily, x='date', y=value_col, title=title, markers=True)
    fig.update_traces(line_color=color)
    return fig

def plot_top10(df, value_col, title_col=None, title="Top10"):
    if value_col not in df:
        return None
    if title_col and title_col in df:
        plot_df = df[[title_col, value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = title_col
    elif 'description' in df:
        plot_df = df[['description', value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = 'description'
    elif 'video_id' in df:
        plot_df = df[['video_id', value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = 'video_id'
    else:
        plot_df = df[[value_col]].dropna().reset_index()
        plot_df['index'] = plot_df['index'] + 1
        plot_df.rename(columns={'index': '视频序号'}, inplace=True)
        x_axis = '视频序号'
    fig = px.bar(plot_df, x=x_axis, y=value_col, title=title, text_auto=True)
    fig.update_layout(xaxis_title="视频", yaxis_title=value_col, xaxis_tickangle=-45)
    return fig

def plot_scatter(df, x_col, y_col, size_col=None, color_col=None, title="相关关系"):
    if x_col not in df or y_col not in df:
        return None
    cols = [x_col, y_col]
    if size_col and size_col in df.columns:
        cols.append(size_col)
    if color_col and color_col in df.columns:
        cols.append(color_col)
    plot_df = df[cols].dropna()
    if plot_df.empty:
        return None
    fig = px.scatter(plot_df, x=x_col, y=y_col,
                     size=size_col if size_col in plot_df.columns else None,
                     color=color_col if color_col in plot_df.columns else None,
                     title=title, opacity=0.6)
    return fig

def plot_author_performance(df):
    if 'author_name' not in df.columns:
        return None
    agg_dict = {'video_id': 'count'}
    if 'like_count' in df:
        agg_dict['like_count'] = 'sum'
    if 'comment_count' in df:
        agg_dict['comment_count'] = 'sum'
    author_stats = df.groupby('author_name').agg(agg_dict).reset_index()
    author_stats.rename(columns={'video_id': 'video_count'}, inplace=True)
    if 'like_count' in df:
        author_stats.rename(columns={'like_count': 'total_likes'}, inplace=True)
        author_stats['avg_likes'] = author_stats['total_likes'] / author_stats['video_count']
    if 'fans_count' in df.columns:
        fans = df.groupby('author_name')['fans_count'].max().reset_index()
        author_stats = author_stats.merge(fans, on='author_name', how='left')
    return author_stats

def plot_publish_hour_heatmap(df, date_col):
    if date_col not in df:
        return None
    temp = df[date_col].dropna()
    if temp.empty:
        return None
    hours = temp.dt.hour
    weekdays = temp.dt.dayofweek
    heat_data = pd.crosstab(weekdays, hours)
    heat_data.index = ['周一','周二','周三','周四','周五','周六','周日']
    fig = px.imshow(heat_data, text_auto=True, aspect="auto",
                    title="发布时段热力图 (星期-小时)",
                    labels=dict(x="小时", y="星期", color="发布数量"))
    return fig

# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown(f"数据文件：`{DATA_FILE.name}`")

    # 加载数据
    raw_df = load_data(DATA_FILE)
    df = standardize_columns(raw_df)

    # 侧边栏筛选
    st.sidebar.header("🔍 数据筛选")
    st.sidebar.markdown(f"**原始视频数**: {len(df)}")

    if 'publish_date' in df.columns:
        valid_dates = df['publish_date'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            start_date, end_date = st.sidebar.date_input(
                "选择日期范围",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            mask = (df['publish_date'].dt.date >= start_date) & (df['publish_date'].dt.date <= end_date)
            df = df[mask]
            st.sidebar.info(f"筛选后视频数: {len(df)}")

    for col in ['like_count', 'comment_count', 'share_count', 'collect_count']:
        if col in df.columns and not df[col].isna().all():
            min_val = float(df[col].min())
            max_val = float(df[col].max())
            if min_val < max_val:
                selected = st.sidebar.slider(f"{col} 范围", min_val, max_val, (min_val, max_val))
                df = df[(df[col] >= selected[0]) & (df[col] <= selected[1])]

    # KPI指标
    kpi = compute_kpis(df)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📹 视频总数", f"{kpi['total_videos']:,}")
    with col2:
        st.metric("❤️ 总点赞数", f"{kpi['total_likes']:,.0f}" if kpi['total_likes'] else "无数据")
    with col3:
        st.metric("💬 总评论数", f"{kpi['total_comments']:,.0f}" if kpi['total_comments'] else "无数据")
    with col4:
        st.metric("⭐ 平均点赞/视频", f"{kpi['avg_likes_per_video']:.1f}" if kpi['avg_likes_per_video'] else "--")

    col5, col6, col7 = st.columns(3)
    with col5:
        st.metric("🔄 总分享数", f"{kpi['total_shares']:,.0f}" if kpi['total_shares'] else "无数据")
    with col6:
        st.metric("🔖 总收藏数", f"{kpi['total_collects']:,.0f}" if kpi['total_collects'] else "无数据")
    with col7:
        st.metric("👥 覆盖粉丝数", f"{kpi['total_fans']:,.0f}" if kpi['total_fans'] else "无数据")

    st.markdown("---")

    # 趋势图
    if 'publish_date' in df and 'like_count' in df:
        st.subheader("📈 每日点赞数趋势")
        fig = plot_trend(df, 'publish_date', 'like_count', "每日点赞总数变化", '#E63946')
        if fig:
            st.plotly_chart(fig, width='stretch')

    # 双列布局
    left, right = st.columns(2)
    with left:
        st.subheader("🏆 点赞数 Top10 视频")
        fig = plot_top10(df, 'like_count', title_col='description', title="点赞量最高的10个视频")
        if fig:
            st.plotly_chart(fig, width='stretch')
    with right:
        st.subheader("📊 点赞 vs 评论 关系")
        fig = plot_scatter(df, 'like_count', 'comment_count',
                           size_col='share_count' if 'share_count' in df else None,
                           title="点赞数与评论数散点图")
        if fig:
            st.plotly_chart(fig, width='stretch')

    # 作者分析
    st.subheader("👤 作者维度分析")
    author_df = plot_author_performance(df)
    if author_df is not None and not author_df.empty:
        st.dataframe(author_df, use_container_width=True)
        if 'total_likes' in author_df.columns:
            fig = px.bar(author_df.sort_values('total_likes', ascending=False).head(10),
                         x='author_name', y='total_likes', title="作者总点赞数 Top10", text_auto=True)
            st.plotly_chart(fig, width='stretch')
    else:
        st.info("未找到作者昵称列")

    # 发布时间分析
    st.subheader("⏰ 发布时间分析")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        if 'publish_date' in df:
            df_week = df['publish_date'].dropna().dt.dayofweek
            if not df_week.empty:
                week_counts = df_week.value_counts().sort_index()
                week_labels = ['周一','周二','周三','周四','周五','周六','周日']
                week_df = pd.DataFrame({'星期': week_labels, '视频数': week_counts.values})
                fig = px.bar(week_df, x='星期', y='视频数', title="按星期发布数量")
                st.plotly_chart(fig, width='stretch')
    with col_h2:
        fig = plot_publish_hour_heatmap(df, 'publish_date')
        if fig:
            st.plotly_chart(fig, width='stretch')

    # 原始数据预览
    st.subheader("📄 原始数据预览")
    display_cols = ['作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数', '创建时间']
    existing = [c for c in display_cols if c in raw_df.columns]
    if existing:
        st.dataframe(raw_df[existing].head(100), use_container_width=True)
    else:
        st.dataframe(raw_df.head(100), use_container_width=True)

    # 侧边栏说明
    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info("看板已部署到 Streamlit Cloud，数据来自仓库中的固定文件。")

if __name__ == "__main__":
    main()