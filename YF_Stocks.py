import os
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

os.chdir('C:/Users/JerryCho0702/Desktop/Python/Stocks')  # 設置工作目錄

ticker_symbols = {  # 定義目標股票代碼和對應的中文股名
    '1519.TW': '華城',
    '2059.TW': '川湖',
    '2330.TW': '台積電',
    '2449.TW': '京元電子',
    '3017.TW': '奇鋐',
    '4129.TW': '聯合',
    '5203.TW': '訊連',
    '2383.TW': '台光電',
    '2368.TW': '金像電',
    '2345.TW': '智邦'
}

end_date = datetime.today().strftime('%Y-%m-%d')  # 設置今天的日期為結束日期

with pd.ExcelWriter('stock_data.xlsx') as writer:  # 創建一個 Excel writer 物件
    for ticker_symbol, stock_name in ticker_symbols.items():
        ticker = yf.Ticker(ticker_symbol)  # 建立股票物件

        try:
            hist_data = ticker.history(start='2020-01-01', end=end_date, interval='1d')  # 獲取歷史股價數據，這裡我們取得從2020年1月1日到今天的數據
            if hist_data.empty:
                raise ValueError(f"No data found for {stock_name} ({ticker_symbol})")
            hist_data.index = hist_data.index.tz_convert(None)  # 移除時區資訊
            hist_data = hist_data.sort_index(ascending=False)  # 將數據按日期從最新到最舊排序
            print(hist_data)
        except Exception as e:
            print(f"獲取 {stock_name} ({ticker_symbol}) 歷史數據時發生錯誤: {e}")
            continue

        # 獲取股票的最新本益比 (P/E)
        try:
            pe_ratio = ticker.info.get('trailingPE', 'N/A')
            hist_data['P/E'] = None  # 初始化 P/E 欄位為 None
            hist_data.at[hist_data.index[0], 'P/E'] = pe_ratio  # 將最新的本益比添加到最新的股價後方
            print(f"{stock_name} ({ticker_symbol}) 的本益比 (P/E): {pe_ratio}")
        except Exception as e:
            print(f"獲取 {stock_name} ({ticker_symbol}) 的本益比 (P/E) 時發生錯誤: {e}")

        try:
            sheet_name = f"{stock_name}_{ticker_symbol.replace('.', '_')}"  # 將歷史數據寫入 Excel
            hist_data.to_excel(writer, sheet_name=sheet_name)
            print(f"{stock_name} ({ticker_symbol}) 數據已成功寫入 Excel！")
        except Exception as e:
            print(f"寫入 {stock_name} ({ticker_symbol}) 的 Excel 文件時發生錯誤: {e}")

        # 畫出股價走勢圖
        try:
            hist_data['Close'].plot(title=f"{stock_name} ({ticker_symbol}) 股價走勢")
            plt.xlabel('日期')
            plt.ylabel('收盤價')
            plt.savefig(f"{ticker_symbol}_price_chart.png")
            plt.close()  # 關閉當前圖表，避免重疊
            print(f"{stock_name} ({ticker_symbol}) 股價走勢圖已成功保存！")
        except Exception as e:
            print(f"畫出 {stock_name} ({ticker_symbol}) 股價走勢圖時發生錯誤: {e}")