import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta


def get_12_month_return(ticker):
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365)
    data = yf.download(ticker, start=start_date, end=end_date, progress=False)
    if len(data) < 2:
        return None
    return (data['Adj Close'].iloc[-1] / data['Adj Close'].iloc[0]) - 1

def original_dual_momentum_strategy(total_asset_value):
    strategy_allocation = total_asset_value

    tickers = ['VOO', 'EFA', 'AGG', 'BIL']
    returns = {}

    for ticker in tickers:
        returns[ticker] = get_12_month_return(ticker)

    if returns['VOO'] is None or returns['EFA'] is None or returns['BIL'] is None or returns['AGG'] is None:
        print("Error: Unable to fetch return data for one or more tickers.")
        return None

    if returns['VOO'] > returns['BIL']:
        if returns['VOO'] > returns['EFA']:
            selected_ticker = 'VOO'
        else:
            selected_ticker = 'EFA'
    else:
        selected_ticker = 'AGG'

    allocation = {
        'VOO': 0,
        'EFA': 0,
        'AGG': 0
    }
    allocation[selected_ticker] = float(strategy_allocation)
    return allocation


def update_excel_with_strategy(excel_file, allocation):
    try:
        # 엑셀 파일의 두 번째 시트 읽기
        df = pd.read_excel(excel_file, sheet_name=1)

        # 각 ETF에 대해 할당량 업데이트
        for ticker, amount in allocation.items():
            if ticker in df['Ticker'].values:
                current_price = yf.Ticker(ticker).info.get('regularMarketPrice', 0)
                if current_price > 0:
                    new_quantity = int(amount / current_price)
                    df.loc[df['Ticker'] == ticker, 'Quantity'] = new_quantity
                    df.loc[df['Ticker'] == ticker, 'Assets'] = amount
            else:
                print(f"Warning: {ticker} not found in the Excel sheet.")

        # 엑셀 파일에 저장
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
            # 기존 두 번째 시트 삭제 (있다면)
            if 'Sheet2' in writer.book.sheetnames:
                idx = writer.book.sheetnames.index('Sheet2')
                writer.book.remove(writer.book.worksheets[idx])

            # 새로운 데이터로 두 번째 시트 작성
            df.to_excel(writer, sheet_name='Sheet2', index=False)

        print("Successfully updated the Excel file with new allocations.")
    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")

