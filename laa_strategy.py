import yfinance as yf
import pandas_datareader as pdr
from datetime import datetime, timedelta


def get_sp500_signal():
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365)  # 1년치 데이터
    sp500 = yf.download('^GSPC', start=start_date, end=end_date, progress=False)

    sp500['200MA'] = sp500['Close'].rolling(window=200).mean()
    current_price = sp500['Close'].iloc[-1]
    current_200ma = sp500['200MA'].iloc[-1]

    return current_price > current_200ma


def get_unemployment_signal():
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365)  # 1년치 데이터
    unemployment_data = pdr.get_data_fred('UNRATE', start=start_date, end=end_date)

    unemployment_data['12MA'] = unemployment_data['UNRATE'].rolling(window=12).mean()
    current_rate = unemployment_data['UNRATE'].iloc[-1]
    current_12ma = unemployment_data['12MA'].iloc[-1]

    return current_rate > current_12ma


def laa_strategy(total_asset_value):
    fixed_assets = ['IWD', 'GLD', 'IEF']
    timing_assets = ['QQQM', 'SHY']

    # 고정 자산 할당 (각 25%)
    allocation = {asset: float(total_asset_value * 0.25) for asset in fixed_assets}

    # 타이밍 자산 할당 (25%)
    sp500_above_200ma = get_sp500_signal()
    unemployment_above_12ma = get_unemployment_signal()

    if sp500_above_200ma and not unemployment_above_12ma:
        allocation['QQQM'] = float(total_asset_value * 0.25)
        allocation['SHY'] = 0.0
    else:
        allocation['QQQM'] = 0.0
        allocation['SHY'] = float(total_asset_value * 0.25)

    return allocation