import pandas as pd
import numpy as np
import statsmodels.formula.api as smf
from sklearn.preprocessing import StandardScaler
from statsmodels.stats.outliers_influence import variance_inflation_factor
import io
import os

# ================= 配置區域 =================
FILE_CONFIG = {
    "input_csv": "實證模型.csv",
    "encoding": "cp950",
    "exclude_std_cols": ['代號', 'TSE 產業別'],
    "rename_map": {
        '資產報酬率(2024全年)': '資產報酬率_2024全年',
        '集團/非集團企業(量化)': '集團_非集團企業_量化'
    }
}

# ================= 核心函式 =================

def preprocess_and_standardize():
    """讀取資料並執行標準化"""
    if not os.path.exists(FILE_CONFIG["input_csv"]):
        print(f"找不到檔案: {FILE_CONFIG['input_csv']}")
        return None

    print(f">>> 讀取資料: {FILE_CONFIG['input_csv']}")
    df = pd.read_csv(FILE_CONFIG["input_csv"], encoding=FILE_CONFIG["encoding"])
    
    # 重新命名欄位
    df.rename(columns=FILE_CONFIG["rename_map"], inplace=True)
    
    # 標準化前先處理數值欄位
    numeric_cols = df.select_dtypes(include=['number']).columns
    numeric_cols = [col for col in numeric_cols if col not in FILE_CONFIG["exclude_std_cols"]]
    
    scaler = StandardScaler()
    df[numeric_cols] = scaler.fit_transform(df[numeric_cols])
    
    df.to_csv("實證模型_標準化.csv", index=False, encoding="utf-8-sig")
    print(">>> 標準化完成，已儲存至 實證模型_標準化.csv")
    return df

def check_multicollinearity(model, output_name):
    """自動計算 OLS 模型的 VIF 值，找出多重共線性的元兇並儲存。"""
    X = model.model.exog
    names = model.model.exog_names
    
    vif_data = pd.DataFrame()
    vif_data["變數名稱"] = names
    vif_data["VIF_數值"] = [variance_inflation_factor(X, i) for i in range(X.shape[1])]
    vif_data = vif_data.sort_values(by="VIF_數值", ascending=False).reset_index(drop=True)
    
    print("\n" + "="*40)
    print(f"        🕵️ VIF 共線性診斷: {output_name}")
    print("="*40)
    print(vif_data.head(10).to_string(index=False))
    print("---------------------------------------")
    
    vif_data.to_html(output_name, encoding='utf-8-sig', index=False)
    return vif_data

def save_model_to_html(model, filename_prefix):
    """將模型結果與主要統計量存成 HTML"""
    try:
        # 儲存係數表 (Table 1)
        summary_html = model.summary().tables[1].as_html()
        df_coef = pd.read_html(io.StringIO(summary_html), header=0, index_col=0)[0]
        df_coef.to_html(f"{filename_prefix}.html", encoding='utf-8-sig')
        
        # 儲存統計摘要 (Table 0)
        summary_main_html = model.summary().tables[0].as_html()
        df_main = pd.read_html(io.StringIO(summary_main_html), header=0, index_col=0)[0]
        # 這裡名稱稍微調整以防與係數表衝突
        df_main.to_html(f"{filename_prefix}_主要統計量.html", encoding='utf-8-sig')
        print(f">>> 報告已儲存: {filename_prefix}")
    except Exception as e:
        print(f"儲存 {filename_prefix} 時發生錯誤: {e}")

# ================= 執行流程 =================

def run_analysis():
    df = preprocess_and_standardize()
    if df is None: return

    # --- 1. 假說 1 (基本模型) ---
    print("\n[執行] 假說 1 迴歸分析...")
    # 根據用戶提供公式，剔除 SC 相關變數
    f1 = ("累計報酬率 ~ 資產總額_自然變數+資產報酬率_2024全年+槓桿比+現金比率+OI+"
          "EBIT+EBITDA+董事會規模+董事會獨立性+董事長兼任總經理+高階主管薪酬長期目標+"
          "集團控制型態_量化+現金充裕度+TESG分數+AI變數")
    model1 = smf.ols(formula=f1, data=df).fit(cov_type='HC3')
    save_model_to_html(model1, "實證模型結果")
    check_multicollinearity(model1, "假說1診斷結果.html")

    # --- 2. 假說 2 (交互作用模型 - 完整版) ---
    print("\n[執行] 假說 2 迴歸分析 (完整版)...")
    f2 = ("累計報酬率 ~ 資產總額_自然變數*集團_非集團企業_量化 + 資產報酬率_2024全年*集團_非集團企業_量化 + "
          "槓桿比*集團_非集團企業_量化 + 現金比率*集團_非集團企業_量化 + OI*集團_非集團企業_量化 + "
          "EBIT*集團_非集團企業_量化 + EBITDA*集團_非集團企業_量化 + 董事會規模*集團_非集團企業_量化 + "
          "董事會獨立性*集團_非集團企業_量化 + 董事長兼任總經理*集團_非集團企業_量化 + "
          "高階主管薪酬長期目標*集團_非集團企業_量化 + 集團控制型態_量化*集團_非集團企業_量化 + "
          "現金充裕度 + TESG分數*集團_非集團企業_量化 + AI變數")
    model2 = smf.ols(formula=f2, data=df).fit(cov_type='HC3')
    check_multicollinearity(model2, "假說2診斷結果.html")

    # --- 3. 假說 2 (修正版 - 剔除嚴重共線性變數) ---
    print("\n[執行] 假說 2 修正版迴歸...")
    f2_revised = ("累計報酬率 ~ 資產總額_自然變數*集團_非集團企業_量化 + 資產報酬率_2024全年*集團_非集團企業_量化 + "
                  "槓桿比*集團_非集團企業_量化 + 現金比率*集團_非集團企業_量化 + 董事會規模*集團_非集團企業_量化 + "
                  "董事會獨立性*集團_非集團企業_量化 + 董事長兼任總經理*集團_非集團企業_量化 + "
                  "高階主管薪酬長期目標*集團_非集團企業_量化 + 集團控制型態_量化*集團_非集團企業_量化 + "
                  "現金充裕度 + TESG分數*集團_非集團企業_量化 + AI變數")
    model2_rev = smf.ols(formula=f2_revised, data=df).fit(cov_type='HC1', use_t=True)
    save_model_to_html(model2_rev, "實證模型2結果_修正後")

    # --- 4. 維度分析 (維度 1, 2, 3) ---
    # 同時執行假說 1 與 假說 2 的維度分組模型
    dims = {
        "維度1": "維度1_事件衝擊",
        "維度2": "維度2_供應鏈",
        "維度3": "維度3_風險管理"
    }

    for key, col in dims.items():
        print(f"\n[執行] {key} 分析...")
        
        # 4a. 維度加法模型 (Hypo 1 式)
        # 注意: 維度1公式使用了 OI 排除其他共線性項
        f_dim = f"累計報酬率 ~ 資產總額_自然變數+資產報酬率_2024全年+槓桿比+現金比率+OI+董事會規模+董事會獨立性+董事長兼任總經理+高階主管薪酬長期目標+集團控制型態_量化+現金充裕度+TESG分數+AI變數 + {col}"
        m_dim = smf.ols(formula=f_dim, data=df).fit(cov_type='HC1', use_t=True)
        save_model_to_html(m_dim, f"實證模型結果_{key}")

        # 4b. 維度交互作用模型 (Hypo 2 式)
        f_dim_h2 = (f"累計報酬率 ~ 資產總額_自然變數*集團_非集團企業_量化 + 資產報酬率_2024全年*集團_非集團企業_量化 + "
                    f"槓桿比*集團_非集團企業_量化 + 現金比率*集團_非集團企業_量化 + 董事會規模*集團_非集團企業_量化 + "
                    f"董事會獨立性*集團_非集團企業_量化 + 董事長兼任總經理*集團_非集團企業_量化 + "
                    f"高階主管薪酬長期目標*集團_非集團企業_量化 + 集團控制型態_量化*集團_非集團企業_量化 + "
                    f"現金充裕度 + TESG分數*集團_非集團企業_量化 + AI變數 + {col}*集團_非集團企業_量化")
        
        # 維度1 使用 HC1 並修正，其他維度使用 HC1 並 use_t
        if key == "維度1":
            m_dim_h2 = smf.ols(formula=f_dim_h2, data=df).fit(cov_type='HC1', use_correction=True)
        else:
            m_dim_h2 = smf.ols(formula=f_dim_h2, data=df).fit(cov_type='HC1', use_t=True)
            
        save_model_to_html(m_dim_h2, f"實證模型結果_假說2_{key}")

    print("\n" + "="*40)
    print(">>> 所有迴歸分析任務已成功執行完畢！")
    print("="*40)

if __name__ == "__main__":
    run_analysis()
