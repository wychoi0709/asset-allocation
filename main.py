import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from dual_momentum_strategy import original_dual_momentum_strategy
from vaa_strategy import vaa_aggressive_strategy
from laa_strategy import laa_strategy
import os
from vaa_strategy import get_return, calculate_momentum_score


EXCEL_FILE = 'asset_allocation_results.xlsx'
SUMMARY_SHEET = 'Portfolio Summary'
DETAILS_SHEET = 'Strategy Details'

# 전략 설명을 전역 변수로 정의
strategy_descriptions = {
    'ODM': {
        'VOO': '미국 대형주',
        'EFA': '미국 외 선진국 주식',
        'AGG': '미국 혼합 채권'
    },
    'VAA': {
        'VOO': '미국 대형주',
        'EFA': '선진국 주식',
        'VWO': '개발 도상국 주식',
        'AGG': '미국 혼합 채권',
        'SHY': '미국 단기 국채',
        'IEF': '미국 중기 국채',
        'LQD': '미국 회사채'
    },
    'LAA': {
        'IWD': '미국 대형주',
        'QQQM': '나스닥',
        'GLD': '금',
        'IEF': '미국 중기 국채',
        'SHY': '미국 단기 국채'
    }
}


def save_to_excel(df, sheet_name, file_name=None):
    today = datetime.now().strftime("%y%m%d")

    if not os.path.exists('result'):
        os.makedirs('result')

    if file_name is None:
        file_name = f"{today}_asset_allocation_results.xlsx"
    else:
        file_name = f"{today}_{file_name}"

    file_path = os.path.join('result', file_name)

    if os.path.exists(file_path):
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            if sheet_name in writer.book.sheetnames:
                idx = writer.book.sheetnames.index(sheet_name)
                writer.book.remove(writer.book.worksheets[idx])
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # 새 파일을 생성할 때 필요한 모든 시트를 만듭니다
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # 'Portfolio Summary' 시트가 없으면 빈 시트로 추가
            if sheet_name != 'Portfolio Summary':
                pd.DataFrame().to_excel(writer, sheet_name='Portfolio Summary', index=False)
            # 'Strategy Details' 시트가 없으면 빈 시트로 추가
            if sheet_name != 'Strategy Details':
                pd.DataFrame().to_excel(writer, sheet_name='Strategy Details', index=False)

    print(f"Successfully saved {sheet_name} to {file_path}")
    return file_path


def get_latest_file(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not files:
        return None
    return max([os.path.join(folder_path, f) for f in files], key=os.path.getmtime)


def get_current_price(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        price_fields = ['regularMarketPrice', 'regularMarketPreviousClose', 'previousClose', 'open',
                        'regularMarketOpen']
        for field in price_fields:
            if field in info and info[field] is not None:
                return info[field]
        print(f"Warning: No valid price found for {ticker}")
        return None
    except Exception as e:
        print(f"Error fetching price for {ticker}: {str(e)}")
        return None


def calculate_current_asset_value():
    try:
        result_folder = 'result'
        latest_file = get_latest_file(result_folder)

        if latest_file is None:
            print("No existing asset allocation file found.")
            return None, 0

        with pd.ExcelFile(latest_file) as xls:
            if 'Strategy Details' not in xls.sheet_names:
                print("'Strategy Details' sheet not found in the latest file.")
                return None, 0

            df = pd.read_excel(xls, sheet_name='Strategy Details')

        # Ticker와 Quantity만 선택
        df = df[['Ticker', 'Quantity']]

        # Ticker별로 Quantity를 합산
        df = df.groupby('Ticker')['Quantity'].sum().reset_index()

        # Quantity가 0보다 큰 행만 선택
        df = df[df['Quantity'] > 0]

        # 현재 가격 가져오기
        df['Current Price'] = df['Ticker'].apply(get_current_price)

        # 현재 가치 계산
        df['Current Value'] = df['Quantity'] * df['Current Price']

        # NaN 값이 있는 행 제거
        df = df.dropna()

        # 총 자산 가치 계산
        total_value = df['Current Value'].sum()

        print(f"Loaded data from: {latest_file}")
        return df, total_value
    except Exception as e:
        print(f"Error calculating asset value: {str(e)}")
        return None, 0

def update_summary_sheet(new_value, allocations):
    try:
        # 기존 요약 정보
        summary_data = {
            '리밸런싱 일자': [datetime.now()],
            '자산 가치': [new_value]
        }

        # 티커별 총 매매량 계산
        total_allocations = {}
        for strategy, allocation in allocations.items():
            for ticker, amount in allocation.items():
                if ticker not in total_allocations:
                    total_allocations[ticker] = 0
                total_allocations[ticker] += amount

        # 티커별 현재 가격 가져오기
        current_prices = {ticker: get_current_price(ticker) for ticker in total_allocations.keys()}

        # 티커별 총 매매량 및 금액 정보 추가
        for ticker, amount in total_allocations.items():
            price = current_prices[ticker]
            quantity = int(amount / price) if price and price > 0 else 0
            summary_data[f'{ticker} 수량'] = [quantity]
            summary_data[f'{ticker} 금액'] = [amount]

        df_summary = pd.DataFrame(summary_data)

        file_path = save_to_excel(df_summary, 'Portfolio Summary')
        print(f"Updated Portfolio Summary with total asset value: ${new_value:.2f}")
        print("Added total quantities and amounts for each ticker.")
        return file_path
    except Exception as e:
        print(f"Error updating Portfolio Summary: {str(e)}")
        return None




def update_strategy_details_sheet(allocations, total_asset_value):
    try:
        data = []
        for strategy, allocation in allocations.items():
            strategy_total = sum(allocation.values())
            for ticker, amount in allocation.items():
                price = get_current_price(ticker)
                quantity = int(amount / price) if price and price > 0 else 0

                row = {
                    'Strategy': strategy,
                    'Description': strategy_descriptions.get(strategy, {}).get(ticker, ''),
                    'Ticker': ticker,
                    'Price': price,
                    'Ratio': (amount / strategy_total) * 100 if strategy_total else 0,
                    'Strategy ratio': (strategy_total / total_asset_value) * 100,
                    'Final ratio': (amount / total_asset_value) * 100,
                    'Assets': amount,
                    'Quantity': quantity
                }

                # VAA 전략에만 수익률과 momentum_score 추가
                if strategy == 'VAA':
                    returns = {
                        '1M': get_return(ticker, 30),
                        '3M': get_return(ticker, 90),
                        '6M': get_return(ticker, 180),
                        '12M': get_return(ticker, 365)
                    }

                    row.update({
                        '1M': returns['1M'] * 100 if returns['1M'] is not None else None,
                        '3M': returns['3M'] * 100 if returns['3M'] is not None else None,
                        '6M': returns['6M'] * 100 if returns['6M'] is not None else None,
                        '12M': returns['12M'] * 100 if returns['12M'] is not None else None,
                        'Score': calculate_momentum_score(ticker)
                    })

                data.append(row)

        df = pd.DataFrame(data)
        file_path = save_to_excel(df, 'Strategy Details')
        print("Successfully updated the Strategy Details sheet.")
        return file_path
    except Exception as e:
        print(f"Error updating Strategy Details: {str(e)}")
        return None


def main():
    asset_df, total_asset_value = calculate_current_asset_value()

    if asset_df is None or total_asset_value == 0:
        print("No existing asset allocation file found or current asset value is zero.")
        initial_investment = float(input("Enter your initial investment amount: $"))
        total_asset_value = initial_investment
    else:
        print("Current Asset Values:")
        print(asset_df)
        print(f"\nTotal Asset Value: ${total_asset_value:.2f}")

    allocations = {
        "ODM": original_dual_momentum_strategy(total_asset_value * 0.333),
        "VAA": vaa_aggressive_strategy(total_asset_value * 0.333),
        "LAA": laa_strategy(total_asset_value * 0.334)
    }

    print("\nStrategy Allocations:")
    for strategy, allocation in allocations.items():
        print(f"{strategy}:")
        print(allocation)

    update_summary_sheet(total_asset_value, allocations)
    file_path = update_strategy_details_sheet(allocations, total_asset_value)

    if file_path:
        print(f"All data has been saved to {file_path}")

if __name__ == "__main__":
    main()
