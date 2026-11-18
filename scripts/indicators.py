# -*- coding: utf-8 -*-
"""
indicators.py — MA/RSI/MACD/Support-Resistance/Smart Money Concepts тооцоолол.
Одоогийн dashboard (docs/index.html) эдгээрийг JavaScript-аар client-side
тооцдог тул энэ модуль ОДООГООР pipeline-д ашиглагдахгүй байна. Ирээдүйд
Python талд (жишээ: server-side алерт, дэлгэрэнгүй тайлан автоматжуулах)
шаардлагатай болвол шууд импортлож ашиглах боломжтой, JS хувилбартай нь
2026 оны 9-р сард тулгаж баталгаажуулсан ижил логиктой.
"""
import numpy as np
import pandas as pd


def sma(series, period):
    return series.rolling(period).mean()


def rsi(closes, period=14):
    delta = closes.diff()
    gain = delta.clip(lower=0).rolling(period).mean()
    loss = (-delta.clip(upper=0)).rolling(period).mean()
    rs = gain / loss
    return 100 - 100 / (1 + rs)


def macd(closes, fast=12, slow=26, signal=9):
    ema_fast = closes.ewm(span=fast, adjust=False).mean()
    ema_slow = closes.ewm(span=slow, adjust=False).mean()
    macd_line = ema_fast - ema_slow
    signal_line = macd_line.ewm(span=signal, adjust=False).mean()
    hist = macd_line - signal_line
    return macd_line, signal_line, hist


def find_sr_levels(df, window=5, lookback=300, n=6, cluster_tol=1.025):
    recent = df.tail(lookback).reset_index(drop=True)
    highs, lows = [], []
    for i in range(window, len(recent) - window):
        seg_h = recent['high'].iloc[i - window:i + window + 1]
        seg_l = recent['low'].iloc[i - window:i + window + 1]
        if recent['high'].iloc[i] == seg_h.max():
            highs.append(recent['high'].iloc[i])
        if recent['low'].iloc[i] == seg_l.min():
            lows.append(recent['low'].iloc[i])
    levels = sorted(highs + lows)
    if not levels:
        return []
    clusters, cur = [], [levels[0]]
    for lv in levels[1:]:
        if lv <= cur[-1] * cluster_tol:
            cur.append(lv)
        else:
            clusters.append(cur)
            cur = [lv]
    clusters.append(cur)
    scored = sorted([(np.mean(c), len(c)) for c in clusters], key=lambda x: -x[1])[:n]
    return sorted([s[0] for s in scored])


def split_support_resistance(levels, last_price):
    supports = sorted([s for s in levels if s < last_price], reverse=True)[:3]
    resistances = sorted([s for s in levels if s > last_price])[:3]
    while len(supports) < 3:
        supports.append(last_price * 0.90 if not supports else supports[-1] * 0.94)
    while len(resistances) < 3:
        resistances.append(last_price * 1.10 if not resistances else resistances[-1] * 1.06)
    return supports, resistances


def detect_swings(highs, lows, k=3):
    swings = []
    n = len(highs)
    for i in range(k, n - k):
        window_h = highs[i - k:i + k + 1]
        window_l = lows[i - k:i + k + 1]
        if highs[i] == max(window_h):
            swings.append({'i': i, 'type': 'H', 'price': highs[i]})
        if lows[i] == min(window_l):
            swings.append({'i': i, 'type': 'L', 'price': lows[i]})
    return swings


def alternate_swings(swings):
    out = []
    for s in swings:
        if not out:
            out.append(s)
            continue
        last = out[-1]
        if last['type'] == s['type']:
            if s['type'] == 'H' and s['price'] > last['price']:
                out[-1] = s
            elif s['type'] == 'L' and s['price'] < last['price']:
                out[-1] = s
        else:
            out.append(s)
    return out


def classify_structure(alt_swings):
    last_h, last_l, trend = None, None, None
    out = []
    for s in alt_swings:
        label, event = None, None
        if s['type'] == 'H':
            if last_h:
                label = 'HH' if s['price'] > last_h['price'] else 'LH'
                if label == 'HH':
                    event = 'CHoCH' if trend == 'down' else 'BOS'
                    trend = 'up'
            last_h = s
        else:
            if last_l:
                label = 'LL' if s['price'] < last_l['price'] else 'HL'
                if label == 'LL':
                    event = 'CHoCH' if trend == 'up' else 'BOS'
                    trend = 'down'
            last_l = s
        out.append({**s, 'label': label, 'event': event})
    return out


def wyckoff_phase(df, n=60):
    n = min(n, len(df) - 1)
    recent = df.tail(n)
    rng_high, rng_low = recent['high'].max(), recent['low'].min()
    last = df['close'].iloc[-1]
    pos = (last - rng_low) / (rng_high - rng_low or 1)
    half = n // 2
    vol_first = recent['volume'].iloc[:half].sum()
    vol_second = recent['volume'].iloc[half:].sum()
    vol_trend = 'нэмэгдэж' if vol_second > vol_first * 1.15 else ('буурч' if vol_second < vol_first * 0.85 else 'тогтвортой')
    range_first = recent['high'].iloc[:half].max() - recent['low'].iloc[:half].min()
    range_second = recent['high'].iloc[half:].max() - recent['low'].iloc[half:].min()
    contracting = range_second < range_first * 0.72
    if pos < 0.35 and contracting:
        phase = 'Хуримтлал (Accumulation)'
    elif pos > 0.65 and contracting:
        phase = 'Түгээлт (Distribution)'
    elif pos > 0.6 and not contracting:
        phase = 'Өсөлтийн үе (Markup)'
    elif pos < 0.4 and not contracting:
        phase = 'Уналтын үе (Markdown)'
    else:
        phase = 'Шилжилтийн үе'
    return phase, pos, vol_trend


def full_analysis(df):
    df = df.copy()
    df['ma7'] = sma(df['close'], 7)
    df['ma25'] = sma(df['close'], 25)
    df['ma99'] = sma(df['close'], 99)
    df['rsi'] = rsi(df['close'])
    df['macd'], df['signal'], df['hist'] = macd(df['close'])

    last = df['close'].iloc[-1]
    levels = find_sr_levels(df)
    supports, resistances = split_support_resistance(levels, last)

    raw_swings = detect_swings(df['high'].tolist(), df['low'].tolist(), k=3)
    alt_swings = alternate_swings(raw_swings)
    structure = classify_structure(alt_swings)
    events = [s for s in structure if s['event']][-4:]
    for e in events:
        e['date'] = str(df['date'].iloc[e['i']])

    phase, pos, vol_trend = wyckoff_phase(df)

    return {
        'last': last, 'last_date': str(df['date'].iloc[-1]),
        'ma7': df['ma7'].iloc[-1], 'ma25': df['ma25'].iloc[-1], 'ma99': df['ma99'].iloc[-1],
        'rsi': df['rsi'].iloc[-1], 'macd': df['macd'].iloc[-1], 'signal': df['signal'].iloc[-1],
        'supports': supports, 'resistances': resistances,
        'bos_choch_events': events,
        'wyckoff_phase': phase, 'wyckoff_pos': pos, 'wyckoff_vol_trend': vol_trend,
    }
