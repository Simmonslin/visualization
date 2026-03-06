import pandas as pd
import numpy as np
from scipy.stats import pearsonr
import os
import sys

# ================= 配置區域 =================
CONFIG = {
    "MAIN_CSV": "Pearson係數.csv",
    "TEST_EXCEL": "pearson係數_進一步測試.xlsx",
    "SHEETS": ["維度1", "維度2", "維度3"],
    "MISSING_CHECK_COLS": ["現金比率", "OI", "現金充裕度"],
    "EXCLUDE_COLS": ["代號", "名稱", "TSE產業名稱", "維度1_事件衝擊"],
    "ENCODINGS": ["cp950", "utf-8", "utf-8-sig"],
}
# ===========================================

def safe_read_csv(file_path):
    """自動嘗試不同編碼讀取 CSV"""
    for enc in CONFIG["ENCODINGS"]:
        try:
            return pd.read_csv(file_path, encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError(f"無法讀取檔案 {file_path}，請檢查編碼格式。")

def calculate_paper_format_corr(df, columns=None):
    """
    優化後的相關係數計算：處理常數項、自動轉型、效能優化
    """
    if columns is not None:
        # 只選取存在的欄位
        valid_cols = [c for c in columns if c in df.columns]
        df = df[valid_cols].copy()
    
    # 強制轉為數值型態 (無法轉換的會變成 NaN)
    df = df.apply(pd.to_numeric, errors='coerce')
    
    # 排除全空或變異數為 0 的欄位 (常數欄位會導致 pearsonr 出錯)
    cols_to_keep = []
    for col in df.columns:
        if df[col].nunique() > 1: # 必須有超過一個不重複的值
            cols_to_keep.append(col)
        else:
            print(f"  [提示] 欄位 '{col}' 變異數為 0 或資料不足，已跳過。")
    
    df = df[cols_to_keep]
    cols = df.columns
    n = len(cols)
    
    # 建立結果矩陣
    formatted_matrix = pd.DataFrame('', columns=cols, index=cols)
    
    # 預先計算相關係數矩陣 (效能較快)，僅作為基礎
    # 顯著性 p-value 仍需透過 scipy 計算
    for i in range(n):
        for j in range(n):
            if i == j:
                formatted_matrix.iloc[i, j] = '1.0000'
            elif i > j:
                col1, col2 = cols[i], cols[j]
                # Pairwise deletion
                mask = ~df[col1].isna() & ~df[col2].isna()
                v1, v2 = df[col1][mask], df[col2][mask]
                
                if len(v1) < 3: # 樣本數過少無法計算顯著性
                    formatted_matrix.iloc[i, j] = 'N/A'
                    continue
                
                try:
                    r, p = pearsonr(v1, v2)
                    sig = '***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.10 else ''
                    formatted_matrix.iloc[i, j] = f"{r:.4f}{sig}"
                except Exception:
                    formatted_matrix.iloc[i, j] = 'Error'
                    
    # 加上 (1), (2)... 前綴
    formatted_matrix.index = [f"({i+1}){col}" for i, col in enumerate(formatted_matrix.index)]
    formatted_matrix.columns = [f"({i+1})" for i in range(n)]
    
    return formatted_matrix

def save_to_excel_safe(df, path):
    """處理檔案被開啟導致無法寫入的問題"""
    try:
        df.to_excel(path)
        print(f"  [成功] 檔案已儲存: {path}")
    except PermissionError:
        print(f"  [錯誤] 無法寫入 {path}。請先關閉該 Excel 檔案再重試！")
    except Exception as e:
        print(f"  [錯誤] 儲存 {path} 時發生非預期錯誤: {e}")

def run_analysis():
    print(">>> 開始執行 Pearson 相關分析優化版")
    
    # 1. 處理主 CSV 檔案
    if os.path.exists(CONFIG["MAIN_CSV"]):
        print(f"\n[步驟 1] 處理 {CONFIG['MAIN_CSV']}...")
        try:
            df_main = safe_read_csv(CONFIG["MAIN_CSV"])
            # 自動排除非數值欄位
            numeric_cols = df_main.select_dtypes(include=[np.number]).columns
            result_main = calculate_paper_format_corr(df_main, numeric_cols)
            save_to_excel_safe(result_main, "correlation_table_主分析.xlsx")
        except Exception as e:
            print(f"  [錯誤] 處理主 CSV 時發生錯誤: {e}")
    else:
        print(f"\n[跳過] 找不到 {CONFIG['MAIN_CSV']}")

    # 2. 處理進一步測試 Excel
    if os.path.exists(CONFIG["TEST_EXCEL"]):
        print(f"\n[步驟 2] 處理 {CONFIG['TEST_EXCEL']}...")
        try:
            xl = pd.ExcelFile(CONFIG["TEST_EXCEL"])
            available_sheets = xl.sheet_names
            
            for sheet in CONFIG["SHEETS"]:
                if sheet not in available_sheets:
                    print(f"  [警告] 分頁 '{sheet}' 不存在於 {CONFIG['TEST_EXCEL']} 中，跳過。")
                    continue
                
                print(f"  正在處理分頁: {sheet}...")
                df_sheet = pd.read_excel(xl, sheet_name=sheet)
                
                # 處理缺失值報告 (僅維度1)
                if sheet == "維度1":
                    missing = df_sheet[df_sheet[CONFIG["MISSING_CHECK_COLS"]].isnull().any(axis=1)]
                    if not missing.empty:
                        save_to_excel_safe(missing, "缺失樣本報告_優化版.xlsx")
                
                # 計算相關係數
                # 自動排除不需要的欄位
                analysis_cols = [c for c in df_sheet.columns if c not in CONFIG["EXCLUDE_COLS"]]
                result_sheet = calculate_paper_format_corr(df_sheet, analysis_cols)
                save_to_excel_safe(result_sheet, f"correlation_table_{sheet}.xlsx")
                
        except Exception as e:
            print(f"  [崩潰] 處理 Excel 時發生錯誤: {e}")
    else:
        print(f"\n[跳過] 找不到 {CONFIG['TEST_EXCEL']}")

    print("\n>>> 分析任務完成！")

if __name__ == "__main__":
    # 確保安裝了必要套件
    try:
        import openpyxl
        import scipy
    except ImportError:
        print("缺少必要套件，請執行: pip install pandas scipy openpyxl")
        sys.exit(1)
        
    run_analysis()
