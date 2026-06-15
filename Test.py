import time
import math
from dataclasses import dataclass

import numpy as np
import pandas as pd
import MetaTrader5 as mt5


# =========================
# Config
# =========================
@dataclass
class Config:
    symbol: str = "XAUUSD"
    timeframe = mt5.TIMEFRAME_M5
    bars: int = 400

    # Indicators
    ema_fast: int = 20
    ema_slow: int = 50
    adx_period: int = 14
    atr_period: int = 14

    # Filters
    adx_min: float = 20.0

    # Swing / BOS
    pivot_left: int = 3
    pivot_right: int = 3
    bos_lookback_swings: int = 30  # how far back we search for last swing

    # Pullback entry
    pullback_atr_frac: float = 0.30  # distance to EMA in ATR fractions

    # Risk / trade
    risk_per_trade: float = 0.005  # 0.5% of balance
    rr: float = 1.8
    atr_sl_buffer: float = 0.20  # extra buffer under swing low in ATR
    magic: int = 260226
    max_spread_points: int = 50  # adjust per symbol

    # Loop
    sleep_sec: int = 10


CFG = Config()


# =========================
# MT5 helpers
# =========================
def mt5_init() -> None:
    if not mt5.initialize():
        raise RuntimeError(f"MT5 initialize failed: {mt5.last_error()}")

    info = mt5.account_info()
    if info is None:
        raise RuntimeError("MT5 account_info() failed (is terminal logged in?)")


def get_rates(symbol: str, timeframe, count: int) -> pd.DataFrame:
    rates = mt5.copy_rates_from_pos(symbol, timeframe, 0, count)
    if rates is None or len(rates) < 100:
        raise RuntimeError(f"copy_rates_from_pos failed: {mt5.last_error()}")
    df = pd.DataFrame(rates)
    df["time"] = pd.to_datetime(df["time"], unit="s")
    return df


def points_per_price(symbol: str) -> float:
    info = mt5.symbol_info(symbol)
    if info is None:
        raise RuntimeError("symbol_info failed")
    return info.point


def get_spread_points(symbol: str) -> int:
    tick = mt5.symbol_info_tick(symbol)
    info = mt5.symbol_info(symbol)
    if tick is None or info is None:
        return 999999
    spread = (tick.ask - tick.bid) / info.point
    return int(spread)


def positions_count(symbol: str, magic: int) -> int:
    pos = mt5.positions_get(symbol=symbol)
    if pos is None:
        return 0
    return sum(1 for p in pos if p.magic == magic)


# =========================
# Indicators (no TA-Lib)
# =========================
def ema(series: pd.Series, period: int) -> pd.Series:
    return series.ewm(span=period, adjust=False).mean()


def true_range(df: pd.DataFrame) -> pd.Series:
    prev_close = df["close"].shift(1)
    tr1 = df["high"] - df["low"]
    tr2 = (df["high"] - prev_close).abs()
    tr3 = (df["low"] - prev_close).abs()
    return pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)


def atr(df: pd.DataFrame, period: int) -> pd.Series:
    tr = true_range(df)
    return tr.rolling(period).mean()


def adx(df: pd.DataFrame, period: int) -> pd.Series:
    # Wilder's ADX (simplified)
    high = df["high"]
    low = df["low"]
    close = df["close"]

    up_move = high.diff()
    down_move = -low.diff()

    plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)

    tr = true_range(df)
    atr_w = tr.rolling(period).mean()

    plus_di = 100 * pd.Series(plus_dm).rolling(period).mean() / atr_w
    minus_di = 100 * pd.Series(minus_dm).rolling(period).mean() / atr_w

    dx = (100 * (plus_di - minus_di).abs() / (plus_di + minus_di)).replace(
        [np.inf, -np.inf], np.nan
    )
    return dx.rolling(period).mean()


# =========================
# Swing (Pivot) detection - confirmed pivots only
# =========================
def pivots_high(df: pd.DataFrame, left: int, right: int) -> pd.Series:
    # pivot high at i if high[i] is max in window [i-left, i+right]
    highs = df["high"].values
    out = np.full(len(df), np.nan)

    for i in range(left, len(df) - right):
        window = highs[i - left : i + right + 1]
        if highs[i] == np.max(window):
            out[i] = highs[i]
    return pd.Series(out, index=df.index)


def pivots_low(df: pd.DataFrame, left: int, right: int) -> pd.Series:
    lows = df["low"].values
    out = np.full(len(df), np.nan)

    for i in range(left, len(df) - right):
        window = lows[i - left : i + right + 1]
        if lows[i] == np.min(window):
            out[i] = lows[i]
    return pd.Series(out, index=df.index)


def last_confirmed_pivot(
    series: pd.Series, upto_index: int, lookback: int
) -> tuple[int | None, float | None]:
    start = max(0, upto_index - lookback)
    window = series.iloc[start : upto_index + 1]
    valid = window.dropna()
    if valid.empty:
        return None, None
    idx = valid.index[-1]
    return int(idx), float(valid.iloc[-1])


# =========================
# Risk / lot sizing
# =========================
def calc_lot_by_risk(
    symbol: str, risk_amount: float, sl_distance_price: float
) -> float:
    info = mt5.symbol_info(symbol)
    if info is None:
        return 0.0

    # tick_value and tick_size help translate price move to money
    tick_value = info.trade_tick_value
    tick_size = info.trade_tick_size

    if tick_value <= 0 or tick_size <= 0 or sl_distance_price <= 0:
        return 0.0

    # money per 1 lot for sl distance:
    money_per_lot = (sl_distance_price / tick_size) * tick_value
    if money_per_lot <= 0:
        return 0.0

    lot = risk_amount / money_per_lot

    # clamp to broker limits
    lot = max(info.volume_min, min(lot, info.volume_max))
    # round to step
    step = info.volume_step
    lot = math.floor(lot / step) * step
    return float(lot)


# =========================
# Order execution
# =========================
def send_order(
    symbol: str, direction: str, lot: float, sl: float, tp: float, magic: int
) -> bool:
    tick = mt5.symbol_info_tick(symbol)
    if tick is None:
        return False

    price = tick.ask if direction == "BUY" else tick.bid
    order_type = mt5.ORDER_TYPE_BUY if direction == "BUY" else mt5.ORDER_TYPE_SELL

    request = {
        "action": mt5.TRADE_ACTION_DEAL,
        "symbol": symbol,
        "volume": lot,
        "type": order_type,
        "price": price,
        "sl": sl,
        "tp": tp,
        "deviation": 20,
        "magic": magic,
        "comment": "BOS_PULLBACK_EA",
        "type_time": mt5.ORDER_TIME_GTC,
        "type_filling": mt5.ORDER_FILLING_FOK,
    }

    result = mt5.order_send(request)
    if result is None:
        print("order_send failed:", mt5.last_error())
        return False

    if result.retcode != mt5.TRADE_RETCODE_DONE:
        print("order_send retcode:", result.retcode, result.comment)
        return False

    print(f"ORDER OK {direction} lot={lot} sl={sl} tp={tp}")
    return True


# =========================
# Strategy logic
# =========================
def strategy_step(cfg: Config) -> None:
    # spread filter
    spread_pts = get_spread_points(cfg.symbol)
    if spread_pts > cfg.max_spread_points:
        print(f"Skip: spread too high ({spread_pts} pts)")
        return

    # one position at a time for safety
    if positions_count(cfg.symbol, cfg.magic) > 0:
        return

    df = get_rates(cfg.symbol, cfg.timeframe, cfg.bars)

    # indicators
    df["ema_fast"] = ema(df["close"], cfg.ema_fast)
    df["ema_slow"] = ema(df["close"], cfg.ema_slow)
    df["atr"] = atr(df, cfg.atr_period)
    df["adx"] = adx(df, cfg.adx_period)

    df["pivh"] = pivots_high(df, cfg.pivot_left, cfg.pivot_right)
    df["pivl"] = pivots_low(df, cfg.pivot_left, cfg.pivot_right)

    # use last closed candle (avoid current forming bar)
    i = len(df) - 2

    close = float(df.loc[i, "close"])
    open_ = float(df.loc[i, "open"])
    ema_fast_i = float(df.loc[i, "ema_fast"])
    ema_slow_i = float(df.loc[i, "ema_slow"])
    atr_i = float(df.loc[i, "atr"])
    adx_i = float(df.loc[i, "adx"])

    if np.isnan(atr_i) or np.isnan(adx_i):
        return

    # regime: trend only
    if adx_i < cfg.adx_min:
        # If you want, we can add Range module later
        return

    trend_up = ema_fast_i > ema_slow_i
    trend_dn = ema_fast_i < ema_slow_i

    # last swings (confirmed)
    _, last_sh = last_confirmed_pivot(df["pivh"], i, cfg.bos_lookback_swings)
    _, last_sl = last_confirmed_pivot(df["pivl"], i, cfg.bos_lookback_swings)

    if last_sh is None or last_sl is None:
        return

    # BOS conditions on closed candle
    bos_up = (close > last_sh) and trend_up
    bos_dn = (close < last_sl) and trend_dn

    # Pullback condition: price near EMA fast (distance <= pullback_atr_frac * ATR)
    dist_to_ema = abs(close - ema_fast_i)
    near_ema = dist_to_ema <= (cfg.pullback_atr_frac * atr_i)

    # Candle confirmation
    bullish = close > open_
    bearish = close < open_

    acc = mt5.account_info()
    if acc is None:
        return

    balance = float(acc.balance)
    risk_amount = balance * cfg.risk_per_trade

    if bos_up and near_ema and bullish:
        # SL under last swing low with ATR buffer
        sl = last_sl - (cfg.atr_sl_buffer * atr_i)
        sl_dist = close - sl
        if sl_dist <= 0:
            return

        tp = close + cfg.rr * sl_dist
        lot = calc_lot_by_risk(cfg.symbol, risk_amount, sl_dist)
        if lot > 0:
            send_order(cfg.symbol, "BUY", lot, sl, tp, cfg.magic)

    elif bos_dn and near_ema and bearish:
        sl = last_sh + (cfg.atr_sl_buffer * atr_i)
        sl_dist = sl - close
        if sl_dist <= 0:
            return

        tp = close - cfg.rr * sl_dist
        lot = calc_lot_by_risk(cfg.symbol, risk_amount, sl_dist)
        if lot > 0:
            send_order(cfg.symbol, "SELL", lot, sl, tp, cfg.magic)


def main():
    mt5_init()
    print("EA started:", CFG)

    while True:
        try:
            strategy_step(CFG)
        except Exception as e:
            print("Error:", e)
        time.sleep(CFG.sleep_sec)


if __name__ == "__main__":
    main()
