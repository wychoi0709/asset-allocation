import yfinance as yf
from datetime import datetime, timedelta


def get_close_price(ticker, date):
    data = yf.download(ticker, start=date - timedelta(days=10), end=date + timedelta(days=1), progress=True)
    if not data.empty:
        return data['Adj Close'].iloc[0]
    else:
        return None


def get_return(ticker, days):
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days)
    end_price = get_close_price(ticker, end_date)
    start_price = get_close_price(ticker, start_date)

    if end_price is None or start_price is None:
        return None

    return (end_price / start_price) - 1


def calculate_momentum_score(ticker):
    returns = {
        30: get_return(ticker, 30),
        90: get_return(ticker, 90),
        180: get_return(ticker, 180),
        365: get_return(ticker, 365)
    }

    if None in returns.values():
        return None

    return (12 * returns[30]) + (4 * returns[90]) + (2 * returns[180]) + returns[365]


def vaa_aggressive_strategy(total_asset_value):
    aggressive_assets = ['VOO', 'EFA', 'VWO', 'AGG']
    defensive_assets = ['LQD', 'IEF', 'SHY']

    aggressive_scores = {asset: calculate_momentum_score(asset) for asset in aggressive_assets}
    defensive_scores = {asset: calculate_momentum_score(asset) for asset in defensive_assets}

    if None in aggressive_scores.values() or None in defensive_scores.values():
        print("Error: Unable to calculate momentum scores for one or more assets.")
        return None

    if all(score >= 0 for score in aggressive_scores.values()):
        selected_asset = max(aggressive_scores, key=aggressive_scores.get)
    else:
        selected_asset = max(defensive_scores, key=defensive_scores.get)

    allocation = {asset: 0 for asset in aggressive_assets + defensive_assets}
    allocation[selected_asset] = float(total_asset_value)

    return allocation