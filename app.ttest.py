import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

# ==========================================
# 0. 頁面設定
# ==========================================
st.set_page_config(page_title="環境監測統計檢定系統 (Pro)", layout="wide")

# ==========================================
# 1. 資料處理核心邏輯 (Data Processing)
# ==========================================

def get_excel_template():
    """產生標準 Excel 範本 (含 MDL 欄位)，使用 openpyxl 引擎"""
    output = io.BytesIO()
    data = {
        '測站': ['測站A', '測站A', '測站A', '測站A', '測站A'],
        '測項': ['重金屬-鉛', '重金屬-鉛', '重金屬-鉛', 'SS', 'SS'],
        '時期': ['施工前', '施工前', '施工期間', '施工前', '施工期間'],
        '數值': ['<0.05', '0.08', 'ND', 15.5, 20.0],
        'MDL':  [0.05, 0.05, 0.05, '', ''],
        '法規下限': ['', '', '', '', ''],
        '法規上限': [0.1, 0.1, 0.1, 50, 50],
        '單位': ['mg/L', 'mg/L', 'mg/L', 'mg/L', 'mg/L']
    }
    df_sample = pd.DataFrame(data)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sample.to_excel(writer, index=False, sheet_name='監測數據')
        worksheet = writer.sheets['監測數據']
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            worksheet.column_dimensions[col].width = 15
    return output.getvalue()

def process_censored_data(row):
    """
    處理含有 < 或 ND 的資料
    邏輯修正：
    1. ND -> 取 MDL 值 (需確保 MDL 為數字)
    2. <數值 -> 取數值
    3. 只有 < 符號 -> 嘗試取 MDL
    """
    val = row['數值']
    
    # 嘗試解析 MDL，若非數字則為 NaN
    try:
        mdl = float(row['MDL'])
    except:
        mdl = np.nan
    
    # 1. 若已經是數字
    if isinstance(val, (int, float)):
        return float(val)
    
    # 轉字串並正規化
    val_str = str(val).strip().upper()
    
    # 2. 處理 "ND"
    if "ND" in val_str:
        if pd.notna(mdl):
            return mdl # 依需求：ND採用MDL
        else:
            return np.nan # 有 ND 沒 MDL -> 無效
            
    # 3. 處理 "<"
    if "<" in val_str:
        try:
            # 情況 A: <0.05 -> 切割出 0.05
            num_text = val_str.replace("<", "").strip()
            if num_text:
                return float(num_text)
            
            # 情況 B: 只有 "<" 符號 -> 嘗試使用 MDL
            elif pd.notna(mdl):
                return mdl
            else:
                return np.nan
        except:
            return np.nan

    # 4. 其他文字轉數字
    try:
        return float(val_str)
    except:
        return np.nan

def perform_stats(df_sub):
    """統計核心邏輯"""
    if df_sub.empty:
        return {'status': 'gray', 'status_text': '無數據', 'p_val': 1.0, 'diff': 0}
        
    group_pre = df_sub[df_sub['時期'] == '施工前']['數值'].dropna().values
    group_dur = df_sub[df_sub['時期'] == '施工期間']['數值'].dropna().values
    
    if len(group_pre) < 2 or len(group_dur) < 2:
        return {'status': 'gray', 'status_text': '數據不足', 'p_val': 1.0, 'diff': 0}

    # Meta data
    lower_limit = df_sub['法規下限'].iloc[0]
    upper_limit = df_sub['法規上限'].iloc[0]
    unit = df_sub['單位'].iloc[0] if pd.notna(df_sub['單位'].iloc[0]) else ""
    item_name = df_sub['測項'].iloc[0]

    mean_pre = np.mean(group_pre)
    mean_dur = np.mean(group_dur)
    diff = mean_dur - mean_pre
    
    # [Bug 4 修正] 檢查是否全為常數 (例如全是 ND 轉換的值)
    # 如果兩組數據完全一樣，或者變異數極小，直接判斷無差異
    if np.array_equal(group_pre, group_dur) or (np.std(group_pre) == 0 and np.std(group_dur) == 0):
        p_val = 1.0
        test_method = "數據無變化 (Constant)"
        is_normal = True # 不重要
    else:
        # 常態性檢定
        try:
            if len(group_pre) < 3 or len(group_dur) < 3:
                is_normal = False
            else:
                _, p_norm_pre = stats.shapiro(group_pre)
                _, p_norm_dur = stats.shapiro(group_dur)
                is_normal = (p_norm_pre > 0.05) and (p_norm_dur > 0.05)
        except:
            is_normal = False

        # 差異檢定
        try:
            if is_normal:
                stat, p_val = stats.ttest_ind(group_pre, group_dur, equal_var=False)
                test_method = "t-test (Welch)"
            else:
                stat, p_val = stats.mannwhitneyu(group_pre, group_dur)
                test_method = "Mann-Whitney U"
        except:
            return {'status': 'gray', 'status_text': '計算錯誤', 'p_val': 1.0}

    # Bootstrap CI
    try:
        # [Bug 4 延伸] 若數據無變化，CI 就是 diff 本身
        if test_method == "數據無變化 (Constant)":
            ci_lower, ci_upper = diff, diff
        else:
            n_boot = 1000
            boot_diffs = []
            for _ in range(n_boot):
                s_pre = np.random.choice(group_pre, len(group_pre), replace=True)
                s_dur = np.random.choice(group_dur, len(group_dur), replace=True)
                boot_diffs.append(np.mean(s_dur) - np.mean(s_pre))
            ci_lower = np.percentile(boot_diffs, 2.5)
            ci_upper = np.percentile(boot_diffs, 97.5)
    except:
        ci_lower, ci_upper = diff, diff

    # 燈號判定
    is_significant = p_val < 0.05
    
    if '溶氧量' in str(item_name) or 'DO' in str(item_name):
        is_worse = diff < 0 
    elif 'pH' in str(item_name):
        is_worse = True 
    else:
        is_worse = diff > 0 

    is_violation = False
    if pd.notna(upper_limit) and mean_dur > upper_limit: is_violation = True
    if pd.notna(lower_limit) and mean_dur < lower_limit: is_violation = True
    
    status = "green"
    status_text = "正常"
    
    if is_violation:
        status = "red"
        status_text = "數值違規/超標"
    elif is_significant and is_worse:
        status = "yellow"
        status_text = "顯著變差 (預警)"
    else:
        status = "green"
        status_text = "無顯著異常"

    return {
        'mean_pre': mean_pre, 'mean_dur': mean_dur, 'diff': diff,
        'p_val': p_val, 'ci_lower': ci_lower, 'ci_upper': ci_upper,
        'test_method': test_method, 'status': status, 'status_text': status_text,
        'lower_limit': lower_limit, 'upper_limit': upper_limit, 'unit': unit
    }

# ==========================================
# 2. Sidebar: 檔案上傳
# ==========================================
st.sidebar.title("📁 資料匯入")

st.sidebar.subheader("1. 下載範本")
st.sidebar.download_button(
    label="📥 下載 Excel 範本 (含MDL)",
    data=get_excel_template(),
    file_name="環境監測數據範本_MDL.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.subheader("2. 上傳資料")
uploaded_file = st.sidebar.file_uploader("請上傳您的監測數據 (xlsx)", type=["xlsx"])

st.sidebar.info("""
**數值處理規則說明：**
1. **ND (未檢出)**：直接採用該列的 `MDL` 值。
2. **< 數值** (如 <0.05)：採用數值的一半 (0.025)。
3. **一般數值**：保持不變。
""")

# ==========================================
# 3. 主畫面邏輯
# ==========================================
st.title("🛡️ 環境監測智能統計系統 (MDL Pro版)")

if uploaded_file is None:
    st.info("👈 請先下載範本，填入數據後上傳。")
else:
    try:
        # 使用 openpyxl 引擎讀取
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # [Bug 1 修正] 去除所有欄位名稱的頭尾空白
        df.columns = df.columns.str.strip()
        
        # 欄位檢查
        required_columns = ['測站', '測項', '時期', '數值']
        if not all(col in df.columns for col in required_columns):
            st.error(f"❌ 缺少必要欄位：{required_columns}。請檢查 Excel 標題列是否有錯字。")
            st.stop()
            
        # 確保 MDL 欄位存在
        if 'MDL' not in df.columns:
            st.warning("⚠️ 未偵測到 `MDL` 欄位，'ND' 數據將被視為無效。")
            df['MDL'] = np.nan
        
        # 確保法規欄位存在
        for col in ['法規下限', '法規上限', '單位']:
            if col not in df.columns: df[col] = np.nan

        df['時期'] = df['時期'].astype(str).str.strip()
        
        # 備份原始數值 (轉字串以免 float 顯示問題)
        df['數值_原始'] = df['數值'].astype(str)

        # 應用資料清洗
        df['數值_清洗後'] = df.apply(process_censored_data, axis=1)
        
        # [Bug 5 修正] 顯示被丟棄的資料細節
        invalid_mask = df['數值_清洗後'].isna()
        n_dropped = invalid_mask.sum()
        
        if n_dropped > 0:
            st.warning(f"⚠️ 有 {n_dropped} 筆資料因無法解析 (如 ND 未填 MDL) 而被略過。")
            with st.expander("點擊查看無效資料清單"):
                st.dataframe(df[invalid_mask][['測站', '測項', '時期', '數值_原始', 'MDL']])
        
        # 寫回數值欄位並移除無效列
        df['數值'] = df['數值_清洗後']
        df = df.dropna(subset=['數值'])

        # --- 開始統計運算 ---
        results = []
        stations = sorted(df['測站'].unique())
        items = sorted(df['測項'].unique())

        progress_bar = st.progress(0)
        total = len(stations) * len(items) if len(stations)*len(items) > 0 else 1
        cnt = 0

        for s in stations:
            for i in items:
                sub_df = df[(df['測站']==s) & (df['測項']==i)]
                if not sub_df.empty:
                    res = perform_stats(sub_df)
                    res['測站'] = s
                    res['測項'] = i
                    results.append(res)
                cnt += 1
                progress_bar.progress(cnt / total)
        progress_bar.empty()
        
        res_df = pd.DataFrame(results)

        if res_df.empty:
            st.warning("沒有產生有效統計結果，請檢查數據。")
            st.stop()

        # Dashboard 顯示
        st.subheader("1. 監測總覽")
        c1, c2, c3, c4 = st.columns(4)
        if 'status' in res_df.columns:
            c1.metric("🔴 違規/超標", len(res_df[res_df['status'] == 'red']))
            c2.metric("🟡 顯著變差", len(res_df[res_df['status'] == 'yellow']))
            c3.metric("🟢 正常/改善", len(res_df[res_df['status'] == 'green']))
            c4.metric("⚪ 數據不足", len(res_df[res_df['status'] == 'gray']))

        st.divider()
        st.subheader("2. 異常偵測矩陣")
        
        status_map = {'red': 2, 'yellow': 1, 'green': 0, 'gray': -1}
        res_df['status_code'] = res_df['status'].map(status_map)
        
        annotations = []
        for idx, row in res_df.iterrows():
            symbol = ""
            if row['status']=='gray': symbol="N/A"
            elif row['p_val']<0.05: symbol="*"
            annotations.append(dict(x=row['測站'], y=row['測項'], text=symbol, showarrow=False,
                                  font=dict(color='white' if row['status'] in ['red','green'] else 'black')))

        fig_h = go.Figure(data=go.Heatmap(
            z=res_df['status_code'], x=res_df['測站'], y=res_df['測項'],
            colorscale=[[0,'#BDC3C7'],[0.25,'#BDC3C7'],[0.25,'#2ECC71'],[0.5,'#2ECC71'],
                        [0.5,'#F1C40F'],[0.75,'#F1C40F'],[0.75,'#E74C3C'],[1,'#E74C3C']],
            zmin=-1, zmax=2, hovertemplate="狀態: %{text}", text=res_df['status_text']
        ))
        fig_h.update_layout(annotations=annotations, height=400)
        st.plotly_chart(fig_h, use_container_width=True)

        st.divider()
        st.subheader("3. 詳細檢定分析")
        c_s1, c_s2 = st.columns(2)
        sel_st = c_s1.selectbox("選擇測站", stations)
        sel_it = c_s2.selectbox("選擇測項", items)
        
        target_df = df[(df['測站']==sel_st) & (df['測項']==sel_it)]
        target_res = res_df[(res_df['測站']==sel_st) & (res_df['測項']==sel_it)]

        if not target_df.empty and not target_res.empty:
            res = target_res.iloc[0]
            if res['status'] == 'gray':
                st.info("數據不足。")
            else:
                fig_est = make_subplots(rows=1, cols=2, column_widths=[0.6, 0.4], 
                                      subplot_titles=(f"{sel_it} 分佈", f"差異估計 ({res['test_method']})"))
                
                colors = {'施工前': 'gray', '施工期間': '#E74C3C' if res['status'] in ['red','yellow'] else '#2ECC71'}
                for p in ['施工前', '施工期間']:
                    sub = target_df[target_df['時期']==p]
                    if not sub.empty:
                        fig_est.add_trace(go.Box(
                            y=sub['數值'], x=sub['時期'], name=p, boxpoints='all',
                            jitter=0.5, pointpos=-1.8, marker=dict(color=colors.get(p)),
                            line=dict(color=colors.get(p)), showlegend=False,
                            text=sub['數值_原始'],
                            hovertemplate="轉化數值: %{y}<br>原始輸入: %{text}"
                        ), row=1, col=1)

                if pd.notna(res['upper_limit']):
                    fig_est.add_hline(y=res['upper_limit'], line_dash="dash", line_color="red", row=1, col=1)

                fig_est.add_hline(y=0, line_color="black", row=1, col=2)
                
                # CI 繪圖 (若 constant 則畫點不畫線)
                if res['test_method'] == "數據無變化 (Constant)":
                     fig_est.add_trace(go.Scatter(
                        x=['差異'], y=[res['diff']], mode='markers', marker=dict(size=12, color='black'),
                        hoverinfo='text', text="數據完全相同，無差異"
                    ), row=1, col=2)
                else:
                    fig_est.add_trace(go.Scatter(
                        x=['差異'], y=[res['diff']], mode='markers', marker=dict(size=12, color='black'),
                        error_y=dict(type='data', array=[res['ci_upper']-res['diff']], 
                                   arrayminus=[res['diff']-res['ci_lower']], thickness=2, width=10, color='black')
                    ), row=1, col=2)
                
                fig_est.update_layout(title_text=f"狀態: {res['status_text']} (P={res['p_val']:.4f})")
                st.plotly_chart(fig_est, use_container_width=True)

    except Exception as e:
        st.error(f"發生未預期的錯誤: {e}")
        st.warning("請檢查 Excel 格式是否正確，或嘗試重新整理頁面。")
    except Exception as e:
        st.error(f"❌ 讀取檔案時發生錯誤：{e}")

        st.warning("請確保您上傳的是有效的 Excel 檔，且格式與範本一致。")

