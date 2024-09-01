import yfinance as yf
from datetime import datetime, timedelta


def get_return(ticker, months):
    end_date = datetime.now()
    start_date = end_date - timedelta(days=months * 30)
    data = yf.download(ticker, start=start_date, end=end_date, progress=False)
    if len(data) < 2:
        return None
    return (data['Adj Close'].iloc[-1] / data['Adj Close'].iloc[0]) - 1


def calculate_momentum_score(ticker):
    returns = {
        1: get_return(ticker, 1),
        3: get_return(ticker, 3),
        6: get_return(ticker, 6),
        12: get_return(ticker, 12)
    }

    if None in returns.values():
        return None

    return (12 * returns[1]) + (4 * returns[3]) + (2 * returns[6]) + returns[12]


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