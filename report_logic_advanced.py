#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GA4 月次レポート自動生成スクリプト（高度解析版）
標準版の全機能に加え、CVR分析・リテンション・コンテンツカテゴリ・曜日分析・Paid Social分析を追加。
"""

import sys
import os
import re
import io
import math
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
import japanize_matplotlib
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
from PIL import Image as PILImage

# ─────────────────────────────────────────────────────
# フォント設定（IPAexGothic: 日本語＋ASCII両対応）
# ─────────────────────────────────────────────────────
FONT_CANDIDATES = [
    os.path.join(os.path.dirname(japanize_matplotlib.__file__), 'fonts', 'ipaexg.ttf'),
    '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
    '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc',
    '/System/Library/Fonts/Hiragino Sans GB.ttc',
]
JP_FONT = None
for font_path in FONT_CANDIDATES:
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont('JP', font_path))
            JP_FONT = 'JP'
            break
        except Exception:
            continue
if JP_FONT is None:
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
        JP_FONT = 'HeiseiKakuGo-W5'
    except Exception:
        JP_FONT = 'Helvetica'

# デザインカラー
COLOR_BG        = colors.HexColor('#EEF3F8')
COLOR_PRIMARY   = colors.HexColor('#1A6FAB')
COLOR_SECONDARY = colors.HexColor('#4DA3D4')
COLOR_DARK      = colors.HexColor('#1E2B3A')
COLOR_GRAY      = colors.HexColor('#8A9BB0')
COLOR_WHITE     = colors.white
COLOR_CARD_BG   = colors.HexColor('#FFFFFF')
COLOR_LIGHT     = colors.HexColor('#F5F8FC')
COLOR_ACCENT    = colors.HexColor('#F4A836')
COLOR_GREEN     = colors.HexColor('#4CAF50')
COLOR_RED       = colors.HexColor('#E74C3C')

# matplotlib カラー
MPL_PRIMARY   = '#1A6FAB'
MPL_SECONDARY = '#4DA3D4'
MPL_ACCENT    = '#F4A836'
MPL_GREEN     = '#4CAF50'
MPL_RED       = '#E74C3C'
MPL_BG        = '#F5F8FC'
MPL_GRID      = '#DDE6EF'

PAGE_W, PAGE_H = A4[1], A4[0]  # landscape A4


# ─────────────────────────────────────────────────────
# CSVパーサー（拡張版: リテンション・国別データ対応）
# ─────────────────────────────────────────────────────
def parse_ga4_csv(filepath: str) -> dict:
    """GA4エクスポートCSVを解析してセクション別データを返す"""
    with open(filepath, 'r', encoding='utf-8') as f:
        raw = f.read()

    # ヘッダメタ情報取得
    account = re.search(r'アカウント:\s*(.+)', raw)
    prop    = re.search(r'プロパティ:\s*(.+)', raw)
    start_d = re.search(r'開始日:\s*(\d{8})', raw)
    end_d   = re.search(r'終了日:\s*(\d{8})', raw)

    meta = {
        'account':  account.group(1).strip() if account else '',
        'property': prop.group(1).strip() if prop else '',
        'start':    start_d.group(1) if start_d else '',
        'end':      end_d.group(1) if end_d else '',
    }

    if meta['end']:
        d = meta['end']
        meta['month_label'] = f"{d[:4]}年{int(d[4:6])}月"
    else:
        meta['month_label'] = ''

    def read_section(header_pattern: str):
        lines = raw.split('\n')
        start_i = None
        col_names = None
        for i, line in enumerate(lines):
            if re.search(header_pattern, line):
                col_names = [c.strip() for c in line.strip().split(',')]
                start_i = i + 1
                break
        if start_i is None:
            return None
        block = []
        for line in lines[start_i:]:
            line = line.strip()
            if not line or line.startswith('#'):
                break
            block.append(line)
        if not block:
            return None
        try:
            return pd.read_csv(
                io.StringIO('\n'.join(block)),
                names=col_names,
                header=None,
            )
        except Exception:
            return None

    # ─ 各セクション取得 ─
    df_active = read_section(r'N 日目,アクティブ ユーザー$')
    df_new    = read_section(r'N 日目,新規ユーザー数')
    df_engage = read_section(r'N 日目,アクティブ ユーザーあたりの平均エンゲージメント時間')
    df_ch_new = read_section(r'ユーザーの最初のメインのチャネル グループ.*新規ユーザー数')
    df_ch_ses = read_section(r'セッションのメインのチャネル グループ.*セッション')
    df_pages  = read_section(r'ページ タイトルとスクリーン クラス,表示回数')
    df_events = read_section(r'イベント名,イベント数')
    df_kev    = read_section(r'イベント名,キーイベント')
    df_trend  = read_section(r'N 日目,30 日,7 日,1 日')
    # 高度解析版で追加
    df_retention = read_section(r'日付,0 週目,1 週目')
    df_country   = read_section(r'国,アクティブ ユーザー')

    # ─ KPI集計 ─
    kpis = {}

    if df_trend is not None and '30 日' in df_trend.columns:
        kpis['active_users'] = int(df_trend['30 日'].iloc[-1])
    elif df_active is not None:
        kpis['active_users'] = int(df_active.iloc[:, 1].sum())
    else:
        kpis['active_users'] = 0

    if df_new is not None:
        kpis['new_users'] = int(df_new.iloc[:, 1].sum())
    elif df_ch_new is not None:
        kpis['new_users'] = int(df_ch_new.iloc[:, 1].sum())
    else:
        kpis['new_users'] = 0

    if df_ch_ses is not None:
        kpis['sessions'] = int(df_ch_ses.iloc[:, 1].sum())
    else:
        kpis['sessions'] = 0

    if df_engage is not None:
        avg_sec = df_engage.iloc[:, 1].mean()
        kpis['avg_engage_sec'] = float(avg_sec)
        m = int(avg_sec // 60)
        s = int(avg_sec % 60)
        kpis['avg_engage_str'] = f"{m}分{s:02d}秒"
    else:
        kpis['avg_engage_sec'] = 0
        kpis['avg_engage_str'] = '-'

    if df_kev is not None:
        kpis['key_events_total'] = int(df_kev.iloc[:, 1].sum())
        kev_dict = {}
        for _, row in df_kev.iterrows():
            kev_dict[row.iloc[0]] = int(row.iloc[1])
        kpis['key_events'] = kev_dict
    else:
        kpis['key_events_total'] = 0
        kpis['key_events'] = {}

    # ── 高度解析用KPI ──

    # CVR（全体）
    if kpis['sessions'] > 0:
        kpis['cvr_total'] = kpis['key_events_total'] / kpis['sessions'] * 100
    else:
        kpis['cvr_total'] = 0

    # 個別CVR
    kpis['cvr_detail'] = {}
    for ev_name, ev_count in kpis['key_events'].items():
        if kpis['sessions'] > 0:
            kpis['cvr_detail'][ev_name] = ev_count / kpis['sessions'] * 100

    # 新規ユーザー比率
    if kpis['active_users'] > 0:
        kpis['new_user_ratio'] = kpis['new_users'] / kpis['active_users'] * 100
    else:
        kpis['new_user_ratio'] = 0

    # フォーム完了率（お問い合わせ → お問い合わせ完了）
    kpis['form_completion_rate'] = 0
    if df_pages is not None:
        col_p = df_pages.columns[0]
        col_v = df_pages.columns[1]
        contact_pv = 0
        contact_done_pv = 0
        for _, row in df_pages.iterrows():
            title = str(row[col_p])
            pv = int(row[col_v])
            if 'お問い合わせ完了' in title:
                contact_done_pv += pv
            elif 'お問い合わせ' in title:
                contact_pv += pv
        kpis['contact_pv'] = contact_pv
        kpis['contact_done_pv'] = contact_done_pv
        if contact_pv > 0:
            kpis['form_completion_rate'] = contact_done_pv / contact_pv * 100

    # エンゲージメント品質指標（イベントデータから算出）
    kpis['page_views'] = 0
    kpis['pv_per_session'] = 0
    kpis['scroll_count'] = 0
    kpis['scroll_rate'] = 0
    kpis['click_count'] = 0
    kpis['click_rate'] = 0
    kpis['form_start'] = 0
    kpis['form_submit'] = 0
    kpis['form_start_to_submit'] = 0
    kpis['file_download'] = 0
    if df_events is not None:
        ev_col = df_events.columns[0]
        ev_val = df_events.columns[1]
        ev_dict = {}
        for _, row in df_events.iterrows():
            ev_dict[str(row[ev_col])] = int(row[ev_val])
        kpis['page_views'] = ev_dict.get('page_view', 0)
        kpis['scroll_count'] = ev_dict.get('scroll', 0)
        kpis['click_count'] = ev_dict.get('click', 0)
        kpis['form_start'] = ev_dict.get('form_start', 0)
        kpis['form_submit'] = ev_dict.get('form_submit', 0)
        kpis['file_download'] = ev_dict.get('file_download', 0)
        if kpis['sessions'] > 0:
            kpis['pv_per_session'] = kpis['page_views'] / kpis['sessions']
        if kpis['page_views'] > 0:
            kpis['scroll_rate'] = kpis['scroll_count'] / kpis['page_views'] * 100
            kpis['click_rate'] = kpis['click_count'] / kpis['page_views'] * 100
        if kpis['form_start'] > 0:
            kpis['form_start_to_submit'] = kpis['form_submit'] / kpis['form_start'] * 100

    # 海外アクセス数
    kpis['non_jp_users'] = 0
    if df_country is not None:
        col_c = df_country.columns[0]
        col_v = df_country.columns[1]
        total_country = int(df_country[col_v].sum())
        jp_users = 0
        for _, row in df_country.iterrows():
            if str(row[col_c]).strip() == 'JP':
                jp_users = int(row[col_v])
        kpis['non_jp_users'] = total_country - jp_users

    # コンテンツカテゴリ分析
    kpis['content_categories'] = {}
    if df_pages is not None:
        col_p = df_pages.columns[0]
        col_v = df_pages.columns[1]
        category_map = {
            '小児矯正': ['小児矯正', 'プレオルソ', '子供の歯並び', '予防矯正', 'お口育て', '歯並びは予防', 'MAM'],
            '小児歯科': ['小児歯科', 'シーラント', 'フッ素', '萌出', '子どもの食べる', 'キシリトール'],
            '一般歯科': ['むし歯治療', 'インプラント', '白い詰め物', '被せ物'],
            '矯正歯科（大人）': ['大人のマウスピース矯正', '大人の予防歯科'],
            '予防歯科': ['予防歯科', '親子の予防', '予防プログラム', 'MIペースト'],
            'ホワイトニング': ['ホワイトニング'],
            '訪問歯科': ['訪問歯科'],
            'マタニティ': ['マタニティ'],
            '口腔育成': ['ことばの教室', '口腔機能', '口腔育成', '舌は上顎', '姿勢と歯並び', '足も口腔'],
            '症例集': ['症例集', '矯正歯科治療'],
            '医院情報': ['院長・スタッフ', '院内ツアー', 'ブログ', 'お知らせ', '地図・診療時間', '料金表', 'スタッフ募集', '初めての方', '保育ルーム', '休診日'],
        }
        for _, row in df_pages.iterrows():
            title = str(row[col_p])
            pv = int(row[col_v])
            categorized = False
            for cat, keywords in category_map.items():
                if any(kw in title for kw in keywords):
                    kpis['content_categories'][cat] = kpis['content_categories'].get(cat, 0) + pv
                    categorized = True
                    break
            if not categorized:
                # トップページ判定
                if 'の歯医者' in title and '｜' not in title and '|' not in title:
                    kpis['content_categories']['トップページ'] = kpis['content_categories'].get('トップページ', 0) + pv
                elif 'お問い合わせ' in title:
                    kpis['content_categories']['お問い合わせ'] = kpis['content_categories'].get('お問い合わせ', 0) + pv

    return {
        'meta': meta,
        'kpis': kpis,
        'df_active': df_active,
        'df_new': df_new,
        'df_engage': df_engage,
        'df_ch_new': df_ch_new,
        'df_ch_ses': df_ch_ses,
        'df_pages':  df_pages,
        'df_events': df_events,
        'df_kev':    df_kev,
        'df_trend':  df_trend,
        'df_retention': df_retention,
        'df_country': df_country,
    }


# ─────────────────────────────────────────────────────
# チャート生成ユーティリティ
# ─────────────────────────────────────────────────────
def fig_to_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def make_daily_line_chart(df_active, df_new, month_label: str, width_inch=9, height_inch=3) -> bytes:
    """日別アクティブ・新規ユーザー折れ線グラフ"""
    if df_active is None and df_new is None:
        return None

    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor=MPL_BG)
    ax.set_facecolor(MPL_BG)

    if df_active is not None and len(df_active) > 0:
        days = df_active.iloc[:, 0].astype(int) + 1
        ax.plot(days, df_active.iloc[:, 1], color=MPL_PRIMARY, linewidth=2.2, marker='o',
                markersize=3, label='アクティブユーザー', zorder=3)
        ax.fill_between(days, df_active.iloc[:, 1], alpha=0.12, color=MPL_PRIMARY)
    elif df_new is not None and len(df_new) > 0:
        days = df_new.iloc[:, 0].astype(int) + 1
    else:
        return None

    if df_new is not None and len(df_new) > 0:
        ax.plot(days, df_new.iloc[:, 1], color=MPL_SECONDARY, linewidth=2, linestyle='--',
                marker='o', markersize=3, label='新規ユーザー', zorder=3)

    ax.set_xlim(1, len(days))
    ax.set_xlabel('日', fontsize=9)
    ax.set_ylabel('ユーザー数', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, color=MPL_GRID, linewidth=0.8, linestyle='-')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=8)
    ax.legend(fontsize=9, framealpha=0, loc='upper right')
    ax.set_title(f'{month_label}  日別ユーザー推移', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_channel_bar_chart(df_ch_ses, width_inch=6, height_inch=3.5) -> bytes:
    """チャネル別セッション横棒グラフ"""
    if df_ch_ses is None:
        return None
    col_ch  = df_ch_ses.columns[0]
    col_val = df_ch_ses.columns[1]
    rename_map = {
        'Organic Search': 'オーガニック検索',
        'Paid Search':    '有料検索(広告)',
        'Paid Social':    '有料SNS(広告)',
        'Direct':         'ダイレクト',
        'Referral':       '参照元',
        'Organic Social': 'オーガニックSNS',
        'Unassigned':     '未割り当て',
        'Cross-network':  'クロスネットワーク',
    }
    df = df_ch_ses.copy()
    df[col_ch] = df[col_ch].map(lambda x: rename_map.get(x, x))
    df = df.sort_values(col_val, ascending=True)
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')
    palette = [MPL_PRIMARY, MPL_SECONDARY, '#5BC0DE', '#F4A836', '#A9D18E', '#D9D9D9', '#C8A2C8', '#FF9999']
    bars = ax.barh(df[col_ch], df[col_val],
                   color=[palette[i % len(palette)] for i in range(len(df))],
                   height=0.55, edgecolor='none')
    for bar, val in zip(bars, df[col_val]):
        ax.text(bar.get_width() + max(df[col_val]) * 0.01, bar.get_y() + bar.get_height() / 2,
                f'{int(val):,}', va='center', fontsize=9, color='#333')
    ax.set_xlim(0, max(df[col_val]) * 1.2)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='x', color=MPL_GRID, linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=9)
    ax.set_title('チャネル別セッション数', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_channel_donut_chart(df_ch_new, width_inch=5, height_inch=4) -> bytes:
    """チャネル別新規ユーザー ドーナツチャート"""
    if df_ch_new is None:
        return None
    col_ch  = df_ch_new.columns[0]
    col_val = df_ch_new.columns[1]
    rename_map = {
        'Organic Search': 'オーガニック\n検索',
        'Paid Search':    '有料検索\n(広告)',
        'Paid Social':    '有料SNS\n(広告)',
        'Direct':         'ダイレクト',
        'Referral':       '参照元',
        'Organic Social': 'SNS',
        'Unassigned':     '未割当',
        'Cross-network':  'その他',
    }
    df = df_ch_new.copy()
    df[col_ch] = df[col_ch].map(lambda x: rename_map.get(x, x))
    palette = [MPL_PRIMARY, MPL_SECONDARY, '#5BC0DE', '#F4A836', '#A9D18E', '#D9D9D9', '#C8A2C8', '#FF9999']
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')
    wedges, texts, autotexts = ax.pie(
        df[col_val], labels=df[col_ch], autopct='%1.1f%%',
        colors=palette[:len(df)],
        wedgeprops=dict(width=0.55, edgecolor='white', linewidth=2),
        startangle=90, pctdistance=0.75,
    )
    for t in texts:
        t.set_fontsize(8)
        t.set_color('#333')
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color('white')
        at.set_fontweight('bold')
    ax.set_title('新規ユーザー獲得チャネル', fontsize=11, pad=12, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_key_events_bar(kev_dict: dict, width_inch=8, height_inch=3.5) -> bytes:
    """キーイベント棒グラフ（汎用化: display_mapにないキーもそのまま表示）"""
    display_map = {
        'WEB予約':       'WEB予約',
        'web予約':       'WEB予約',
        '成人電話':       '電話（成人）',
        '小児電話':       '電話（小児）',
        'スマホ_TEL':    '電話タップ',
        '問診票_成人':    '問診票（成人）',
        '問診票_小児':    '問診票（小児）',
    }
    items = {}
    for k, v in kev_dict.items():
        if v > 0:
            label = display_map.get(k, k)
            items[label] = items.get(label, 0) + v
    if not items:
        return None

    sorted_pairs = sorted(items.items(), key=lambda x: x[1], reverse=True)
    labels = [p[0] for p in sorted_pairs]
    values = [p[1] for p in sorted_pairs]

    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')
    palette = [MPL_PRIMARY if i == 0 else MPL_SECONDARY if i == 1 else '#5BC0DE' for i in range(len(labels))]
    bars = ax.bar(labels, values, color=palette, width=0.55, edgecolor='none')
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.02,
                f'{int(val):,}', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#1E2B3A')
    ax.set_ylim(0, max(values) * 1.25)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='y', color=MPL_GRID, linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(axis='x', colors='#555', labelsize=9, rotation=15)
    ax.tick_params(axis='y', colors='#555', labelsize=9)
    ax.set_title('キーイベント（コンバージョン）内訳', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_top_pages_chart(df_pages, top_n=20, width_inch=8, height_inch=6) -> bytes:
    """上位ページ横棒グラフ"""
    if df_pages is None:
        return None
    col_p = df_pages.columns[0]
    col_v = df_pages.columns[1]
    df = df_pages.head(top_n).copy()

    def shorten(title, max_len=28):
        title = str(title)
        # 「川崎市で「XXX」ひらの歯科クリニック」→「XXX」を抽出
        m = re.search(r'[「『](.*?)[」』]', title)
        if m:
            s = m.group(1)
            return s[:max_len] + '…' if len(s) > max_len else s
        # ｜区切りで分割
        parts = [p.strip() for p in re.split(r'[｜|]', title) if p.strip()]
        # 医院名・地域名パーツを除外
        noise = r'(歯医者|歯科クリニック|歯科医院|武蔵新城|新丸子|川崎市|北本市)'
        meaningful = [p for p in parts if not re.search(noise, p)]
        if meaningful:
            s = meaningful[0]
        elif len(parts) >= 2:
            s = 'トップページ'
        else:
            s = parts[-1] if parts else title
        # 冒頭の地名除去
        s = re.sub(r'^(川崎市で|北本市で)', '', s).strip()
        return s[:max_len] + '…' if len(s) > max_len else s

    df['label'] = df[col_p].apply(shorten)
    df = df.sort_values(col_v, ascending=True)
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')
    palette = [MPL_PRIMARY if i == len(df) - 1 else MPL_SECONDARY if i >= len(df) - 5 else '#A8C7E2'
               for i in range(len(df))]
    bars = ax.barh(df['label'], df[col_v], color=palette, height=0.6, edgecolor='none')
    for bar, val in zip(bars, df[col_v]):
        ax.text(bar.get_width() + max(df[col_v]) * 0.01, bar.get_y() + bar.get_height() / 2,
                f'{int(val):,}', va='center', fontsize=8, color='#333')
    ax.set_xlim(0, max(df[col_v]) * 1.18)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='x', color=MPL_GRID, linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=8)
    ax.set_title(f'上位{top_n}ページ（表示回数）', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_breakdown_bar_chart(data_dict: dict, title: str, width_inch=6, height_inch=3) -> bytes:
    """内訳用小規模横棒グラフ"""
    if not data_dict:
        return None
    labels = list(data_dict.keys())
    values = list(data_dict.values())
    sorted_pairs = sorted(zip(values, labels), reverse=False)
    values, labels = zip(*sorted_pairs)
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')
    palette = [MPL_PRIMARY, MPL_SECONDARY, '#5BC0DE', '#F4A836', '#A9D18E']
    bars = ax.barh(labels, values,
                   color=[palette[i % len(palette)] for i in range(len(labels))],
                   height=0.6, edgecolor='none')
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.02, bar.get_y() + bar.get_height() / 2,
                f'{int(val):,}', va='center', fontsize=9, fontweight='bold', color='#333')
    ax.set_xlim(0, max(values) * 1.3)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='x', color=MPL_GRID, linewidth=0.5, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=9)
    ax.set_title(title, fontsize=10, pad=6, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


# ─────────────────────────────────────────────────────
# 高度解析版の追加チャート
# ─────────────────────────────────────────────────────
def make_cvr_gauge_chart(cvr_total, cvr_detail, sessions, width_inch=9, height_inch=3.5) -> bytes:
    """CVR分析チャート: 全体CVR + 個別CVR横棒"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(width_inch, height_inch),
                                    facecolor='white', gridspec_kw={'width_ratios': [1, 2]})

    # 左: 全体CVRゲージ風
    ax1.set_facecolor('white')
    ax1.set_xlim(0, 1)
    ax1.set_ylim(0, 1)
    ax1.text(0.5, 0.65, f'{cvr_total:.1f}%', fontsize=36, fontweight='bold',
             ha='center', va='center', color=MPL_PRIMARY)
    ax1.text(0.5, 0.4, '全体CVR', fontsize=12, ha='center', va='center', color='#555')
    ax1.text(0.5, 0.25, f'({sessions:,}セッション中)', fontsize=9,
             ha='center', va='center', color='#999')
    ax1.axis('off')

    # 右: 個別CVR横棒
    ax2.set_facecolor('white')
    display_map = {
        'web予約': 'WEB予約', 'WEB予約': 'WEB予約',
        'スマホ_TEL': '電話タップ', '成人電話': '電話（成人）', '小児電話': '電話（小児）',
    }
    items = {}
    for k, v in cvr_detail.items():
        label = display_map.get(k, k)
        items[label] = items.get(label, 0) + v

    if items:
        sorted_items = sorted(items.items(), key=lambda x: x[1], reverse=False)
        labels = [x[0] for x in sorted_items]
        values = [x[1] for x in sorted_items]
        palette = [MPL_PRIMARY, MPL_SECONDARY, '#5BC0DE', '#F4A836']
        bars = ax2.barh(labels, values,
                       color=[palette[i % len(palette)] for i in range(len(labels))],
                       height=0.5, edgecolor='none')
        for bar, val in zip(bars, values):
            ax2.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height() / 2,
                    f'{val:.1f}%', va='center', fontsize=10, fontweight='bold', color='#333')
        ax2.set_xlim(0, max(values) * 1.4)
        ax2.set_title('イベント別コンバージョン率', fontsize=11, pad=8, color='#1E2B3A')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['left'].set_color(MPL_GRID)
    ax2.spines['bottom'].set_color(MPL_GRID)
    ax2.grid(True, axis='x', color=MPL_GRID, linewidth=0.5, linestyle='--')
    ax2.tick_params(colors='#555', labelsize=9)
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_retention_heatmap(df_retention, width_inch=9, height_inch=3.5) -> bytes:
    """リテンション（ユーザー維持率）ヒートマップ"""
    if df_retention is None:
        return None

    df = df_retention.copy()
    date_col = df.columns[0]
    week_cols = [c for c in df.columns[1:] if c != date_col]

    # -1を0に置換
    for col in week_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df[col] = df[col].apply(lambda x: 0 if x < 0 else x)

    # 維持率に変換（0週目を基準）
    rate_data = []
    labels_y = df[date_col].tolist()
    for _, row in df.iterrows():
        base = row[week_cols[0]]
        if base > 0:
            rates = [row[c] / base * 100 if row[c] > 0 else 0 for c in week_cols]
        else:
            rates = [0] * len(week_cols)
        rate_data.append(rates)

    rate_arr = np.array(rate_data)

    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')

    cmap = plt.cm.Blues
    im = ax.imshow(rate_arr, cmap=cmap, aspect='auto', vmin=0, vmax=100)

    ax.set_xticks(range(len(week_cols)))
    ax.set_xticklabels(week_cols, fontsize=8)
    ax.set_yticks(range(len(labels_y)))
    ax.set_yticklabels([str(l)[:12] for l in labels_y], fontsize=8)

    # セル内にテキスト
    for i in range(len(labels_y)):
        for j in range(len(week_cols)):
            val = rate_arr[i, j]
            if val > 0:
                txt_color = 'white' if val > 50 else '#333'
                if j == 0:
                    # 0週目は実数表示
                    ax.text(j, i, f'{int(df.iloc[i][week_cols[j]]):,}',
                            ha='center', va='center', fontsize=8, color=txt_color)
                else:
                    ax.text(j, i, f'{val:.1f}%',
                            ha='center', va='center', fontsize=8, color=txt_color)

    ax.set_title('ユーザー維持率（週別コホート）', fontsize=11, pad=8, color='#1E2B3A')
    fig.colorbar(im, ax=ax, label='維持率 (%)', shrink=0.8)
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_content_category_chart(categories: dict, width_inch=9, height_inch=4) -> bytes:
    """コンテンツカテゴリ分析（円グラフ + 横棒グラフ）"""
    if not categories:
        return None

    # 小さすぎるカテゴリを「その他」に統合
    total = sum(categories.values())
    main_cats = {}
    other = 0
    for k, v in sorted(categories.items(), key=lambda x: x[1], reverse=True):
        if v / total >= 0.02:
            main_cats[k] = v
        else:
            other += v
    if other > 0:
        main_cats['その他'] = other

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(width_inch, height_inch),
                                    facecolor='white', gridspec_kw={'width_ratios': [1, 1.2]})

    # 左: ドーナツ
    palette = [MPL_PRIMARY, MPL_SECONDARY, '#5BC0DE', '#F4A836', '#A9D18E',
               '#C8A2C8', '#FF9999', '#D9D9D9', '#87CEEB', '#FFD700', '#90EE90']
    ax1.set_facecolor('white')
    sorted_cats = sorted(main_cats.items(), key=lambda x: x[1], reverse=True)
    labels_pie = [x[0] for x in sorted_cats]
    values_pie = [x[1] for x in sorted_cats]
    wedges, texts, autotexts = ax1.pie(
        values_pie, labels=labels_pie, autopct='%1.0f%%',
        colors=palette[:len(labels_pie)],
        wedgeprops=dict(width=0.5, edgecolor='white', linewidth=1.5),
        startangle=90, pctdistance=0.78,
    )
    for t in texts:
        t.set_fontsize(7)
        t.set_color('#333')
    for at in autotexts:
        at.set_fontsize(7)
        at.set_color('white')
        at.set_fontweight('bold')
    ax1.set_title('カテゴリ別PV比率', fontsize=10, pad=8, color='#1E2B3A')

    # 右: 横棒
    ax2.set_facecolor('white')
    sorted_bar = sorted(main_cats.items(), key=lambda x: x[1], reverse=False)
    bar_labels = [x[0] for x in sorted_bar]
    bar_values = [x[1] for x in sorted_bar]
    bars = ax2.barh(bar_labels, bar_values,
                   color=[palette[i % len(palette)] for i in range(len(bar_labels))],
                   height=0.6, edgecolor='none')
    for bar, val in zip(bars, bar_values):
        ax2.text(bar.get_width() + max(bar_values) * 0.02, bar.get_y() + bar.get_height() / 2,
                f'{int(val):,} PV', va='center', fontsize=8, color='#333')
    ax2.set_xlim(0, max(bar_values) * 1.25)
    ax2.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax2.grid(True, axis='x', color=MPL_GRID, linewidth=0.5, linestyle='--')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['left'].set_color(MPL_GRID)
    ax2.spines['bottom'].set_color(MPL_GRID)
    ax2.tick_params(colors='#555', labelsize=8)
    ax2.set_title('カテゴリ別 表示回数', fontsize=10, pad=8, color='#1E2B3A')

    fig.tight_layout()
    return fig_to_bytes(fig)


def make_weekday_chart(df_active, df_new, start_date_str: str, width_inch=9, height_inch=3.5) -> bytes:
    """曜日別パターン分析"""
    if df_active is None:
        return None

    # 開始日から曜日を割り当て
    try:
        start_date = datetime.strptime(start_date_str, '%Y%m%d')
    except (ValueError, TypeError):
        return None

    days_data = df_active.iloc[:, 1].values
    weekday_names = ['月', '火', '水', '木', '金', '土', '日']
    weekday_totals = {d: [] for d in weekday_names}

    for i, val in enumerate(days_data):
        dt = start_date + timedelta(days=i)
        wd = weekday_names[dt.weekday()]
        weekday_totals[wd].append(val)

    avg_by_weekday = {wd: np.mean(vals) for wd, vals in weekday_totals.items() if vals}

    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')

    labels = weekday_names
    values = [avg_by_weekday.get(d, 0) for d in labels]
    overall_avg = np.mean(values) if values else 0

    bar_colors = [MPL_PRIMARY if v >= overall_avg else MPL_SECONDARY for v in values]
    bars = ax.bar(labels, values, color=bar_colors, width=0.6, edgecolor='none')

    # 平均線
    ax.axhline(y=overall_avg, color=MPL_ACCENT, linestyle='--', linewidth=1.5, label=f'全体平均: {overall_avg:.0f}')

    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.02,
                f'{val:.0f}', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#1E2B3A')

    ax.set_ylim(0, max(values) * 1.25)
    ax.set_ylabel('平均アクティブユーザー数', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='y', color=MPL_GRID, linewidth=0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=10)
    ax.legend(fontsize=9, framealpha=0)
    ax.set_title('曜日別 平均アクティブユーザー数', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_new_vs_returning_chart(new_users, active_users, width_inch=4, height_inch=3.5) -> bytes:
    """新規 vs リピーター ドーナツチャート"""
    returning = max(0, active_users - new_users)
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')

    values = [new_users, returning]
    labels = [f'新規\n{new_users:,}人', f'リピーター\n{returning:,}人']
    chart_colors = [MPL_SECONDARY, MPL_PRIMARY]

    wedges, texts, autotexts = ax.pie(
        values, labels=labels, autopct='%1.1f%%',
        colors=chart_colors,
        wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2),
        startangle=90, pctdistance=0.75,
    )
    for t in texts:
        t.set_fontsize(9)
        t.set_color('#333')
    for at in autotexts:
        at.set_fontsize(9)
        at.set_color('white')
        at.set_fontweight('bold')

    ax.set_title('新規 vs リピーター', fontsize=11, pad=10, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_paid_social_chart(df_ch_ses, df_ch_new, width_inch=9, height_inch=3.5) -> bytes:
    """Paid Social効果分析チャート"""
    if df_ch_ses is None or df_ch_new is None:
        return None

    col_ch_s = df_ch_ses.columns[0]
    col_val_s = df_ch_ses.columns[1]
    col_ch_n = df_ch_new.columns[0]
    col_val_n = df_ch_new.columns[1]

    # Paid Socialのデータ抽出
    paid_social_ses = df_ch_ses[df_ch_ses[col_ch_s] == 'Paid Social']
    paid_social_new = df_ch_new[df_ch_new[col_ch_n] == 'Paid Social']

    if paid_social_ses.empty and paid_social_new.empty:
        return None

    ps_sessions = int(paid_social_ses[col_val_s].iloc[0]) if not paid_social_ses.empty else 0
    ps_new_users = int(paid_social_new[col_val_n].iloc[0]) if not paid_social_new.empty else 0
    total_sessions = int(df_ch_ses[col_val_s].sum())
    total_new = int(df_ch_new[col_val_n].sum())

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(width_inch, height_inch), facecolor='white')

    # 左: セッション内のPaid Social割合
    ax1.set_facecolor('white')
    other_ses = total_sessions - ps_sessions
    ax1.pie([ps_sessions, other_ses],
            labels=[f'Paid Social\n{ps_sessions:,}', f'その他\n{other_ses:,}'],
            autopct='%1.1f%%', colors=[MPL_ACCENT, '#DDE6EF'],
            wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2),
            startangle=90, pctdistance=0.75,
            textprops={'fontsize': 9})
    ax1.set_title('セッション構成比', fontsize=10, pad=8, color='#1E2B3A')

    # 右: 新規ユーザー内のPaid Social割合
    ax2.set_facecolor('white')
    other_new = total_new - ps_new_users
    ax2.pie([ps_new_users, other_new],
            labels=[f'Paid Social\n{ps_new_users:,}', f'その他\n{other_new:,}'],
            autopct='%1.1f%%', colors=[MPL_ACCENT, '#DDE6EF'],
            wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2),
            startangle=90, pctdistance=0.75,
            textprops={'fontsize': 9})
    ax2.set_title('新規ユーザー構成比', fontsize=10, pad=8, color='#1E2B3A')

    fig.suptitle('Paid Social（有料SNS広告）効果分析', fontsize=12, y=1.02, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


def make_form_funnel_chart(contact_pv, contact_done_pv, completion_rate, width_inch=5, height_inch=3.5) -> bytes:
    """フォーム完了率ファネルチャート"""
    fig, ax = plt.subplots(figsize=(width_inch, height_inch), facecolor='white')
    ax.set_facecolor('white')

    stages = ['お問い合わせページ\n閲覧', 'お問い合わせ\n完了']
    values = [contact_pv, contact_done_pv]
    bar_colors = [MPL_SECONDARY, MPL_PRIMARY]

    bars = ax.bar(stages, values, color=bar_colors, width=0.5, edgecolor='none')
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.03,
                f'{int(val):,}', ha='center', va='bottom', fontsize=14, fontweight='bold', color='#1E2B3A')

    # 完了率の矢印テキスト
    mid_x = 0.5
    mid_y = (contact_pv + contact_done_pv) / 2
    ax.annotate(f'完了率\n{completion_rate:.1f}%',
                xy=(1, contact_done_pv), xytext=(0.5, mid_y),
                fontsize=12, fontweight='bold', color=MPL_ACCENT,
                ha='center', va='center',
                arrowprops=dict(arrowstyle='->', color=MPL_ACCENT, lw=2))

    ax.set_ylim(0, max(values) * 1.35)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, axis='y', color=MPL_GRID, linewidth=0.5, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(MPL_GRID)
    ax.spines['bottom'].set_color(MPL_GRID)
    ax.tick_params(colors='#555', labelsize=10)
    ax.set_title('お問い合わせフォーム完了率', fontsize=11, pad=8, color='#1E2B3A')
    fig.tight_layout()
    return fig_to_bytes(fig)


# ─────────────────────────────────────────────────────
# PDF生成（ReportCanvas）
# ─────────────────────────────────────────────────────
class ReportCanvas:
    """カスタムページレイアウト（reportlab canvas直接描画）"""

    def __init__(self, filepath: str):
        self.c = canvas.Canvas(filepath, pagesize=(PAGE_W, PAGE_H))
        self.c.setTitle('GA4月次レポート（高度解析版）')

    def save(self):
        self.c.save()

    def new_page(self):
        self.c.showPage()

    def _bg(self, color=None):
        clr = color or COLOR_BG
        self.c.setFillColor(clr)
        self.c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    def _text(self, x, y, text, font_size=10, color=COLOR_DARK, bold=False, align='left'):
        self.c.setFillColor(color)
        self.c.setFont(JP_FONT, font_size)
        if align == 'center':
            self.c.drawCentredString(x, y, str(text))
        elif align == 'right':
            self.c.drawRightString(x, y, str(text))
        else:
            self.c.drawString(x, y, str(text))

    def _rect(self, x, y, w, h, fill_color=COLOR_PRIMARY, stroke_color=None, radius=3):
        self.c.setFillColor(fill_color)
        if stroke_color:
            self.c.setStrokeColor(stroke_color)
            self.c.roundRect(x, y, w, h, radius, fill=1, stroke=1)
        else:
            self.c.setStrokeColor(fill_color)
            self.c.roundRect(x, y, w, h, radius, fill=1, stroke=0)

    def _line(self, x1, y1, x2, y2, color=COLOR_GRAY, width=0.5):
        self.c.setStrokeColor(color)
        self.c.setLineWidth(width)
        self.c.line(x1, y1, x2, y2)

    def _image_bytes(self, img_bytes, x, y, w, h=None):
        from reportlab.lib.utils import ImageReader
        buf = io.BytesIO(img_bytes)
        pil_img = PILImage.open(buf)
        pw, ph = pil_img.size
        aspect = ph / pw
        actual_h = h if h else w * aspect
        buf.seek(0)
        ir = ImageReader(buf)
        self.c.drawImage(ir, x, y - actual_h, width=w, height=actual_h,
                         preserveAspectRatio=True, mask='auto')

    def _footer(self, company_name='YOUR COMPANY', page_num=1, client_name=''):
        self._line(10 * mm, 10 * mm, PAGE_W - 10 * mm, 10 * mm,
                   color=COLOR_GRAY, width=0.4)
        self._text(12 * mm, 7 * mm, company_name, font_size=7, color=COLOR_GRAY)
        self._text(PAGE_W - 12 * mm, 7 * mm, str(page_num), font_size=7,
                   color=COLOR_GRAY, align='right')

    def _section_header(self, x, y, title: str, subtitle: str = ''):
        self._rect(x, y - 1 * mm, 3.5 * mm, 7 * mm, fill_color=COLOR_PRIMARY)
        self._text(x + 5 * mm, y + 3 * mm, title, font_size=13, color=COLOR_DARK)
        if subtitle:
            self._text(x + 5 * mm, y - 1 * mm, subtitle, font_size=8, color=COLOR_GRAY)

    def _page_title_bar(self, title: str, subtitle: str = ''):
        self._rect(0, PAGE_H - 14 * mm, PAGE_W, 14 * mm, fill_color=COLOR_PRIMARY)
        self._text(12 * mm, PAGE_H - 9 * mm, title, font_size=14, color=COLOR_WHITE)
        if subtitle:
            self._text(PAGE_W - 12 * mm, PAGE_H - 9 * mm,
                       subtitle, font_size=9, color=COLOR_WHITE, align='right')

    # ─── P1: 表紙 ───
    def draw_cover(self, meta, company_name='YOUR COMPANY', department='', logo_path=None):
        self._bg(COLOR_BG)
        if logo_path and os.path.exists(logo_path):
            self.c.drawImage(logo_path, 15 * mm, PAGE_H - 25 * mm,
                             width=52.5 * mm, height=15 * mm, preserveAspectRatio=True, mask='auto')

        client_label = meta.get('property', '').replace(' - GA4', '')
        self._text(27 * mm, PAGE_H * 0.62, client_label + '　御中', font_size=27, color=COLOR_DARK)
        self._text(27 * mm, PAGE_H * 0.48, 'GA4 月次レポート', font_size=45, color=COLOR_DARK)
        self._text(27 * mm, PAGE_H * 0.42, '（詳細版）', font_size=18, color=COLOR_SECONDARY)
        self._text(27 * mm, PAGE_H * 0.34, meta.get('month_label', ''), font_size=24, color=COLOR_GRAY)
        self._line(27 * mm, PAGE_H * 0.31, 150 * mm, PAGE_H * 0.31, color=COLOR_PRIMARY, width=2.0)

        bottom_x = PAGE_W - 20 * mm
        self._text(bottom_x, PAGE_H * 0.2 + 9 * mm, company_name, font_size=15, color=COLOR_DARK, align='right')
        if department:
            self._text(bottom_x, PAGE_H * 0.2, department, font_size=15, color=COLOR_DARK, align='right')
        self._footer(company_name, 1, client_label)

    # ─── P2: サマリ ───
    def draw_summary(self, meta, kpis, daily_chart_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        period = ''
        if meta.get('start') and meta.get('end'):
            s, e = meta['start'], meta['end']
            period = f"{s[:4]}/{s[4:6]}/{s[6:]} ─ {e[:4]}/{e[4:6]}/{e[6:]}"
        self._page_title_bar('GA4 Summary', period)

        # KPIカード（6枚: 従来4 + CVR + 新規率）
        card_y = PAGE_H - 22 * mm
        card_h = 28 * mm
        card_labels = ['月間アクティブユーザー', '新規ユーザー', 'セッション数',
                       '平均エンゲージメント時間', '全体CVR', '新規ユーザー比率']
        card_values = [
            f"{kpis['active_users']:,}", f"{kpis['new_users']:,}",
            f"{kpis['sessions']:,}", kpis['avg_engage_str'],
            f"{kpis['cvr_total']:.1f}%", f"{kpis['new_user_ratio']:.0f}%",
        ]
        card_units = ['人', '人', '件', '', '', '']

        num_cards = 6
        margin = 10 * mm
        total_gap = (num_cards - 1) * 3 * mm
        card_w = (PAGE_W - 2 * margin - total_gap) / num_cards

        for i, (lbl, val, unit) in enumerate(zip(card_labels, card_values, card_units)):
            cx = margin + i * (card_w + 3 * mm)
            self.c.setFillColor(COLOR_CARD_BG)
            self.c.setStrokeColor(colors.HexColor('#DDE6EF'))
            self.c.roundRect(cx, card_y - card_h, card_w, card_h, 4, fill=1, stroke=1)
            self._rect(cx, card_y - 3 * mm, card_w, 3 * mm, fill_color=COLOR_PRIMARY, radius=0)
            self._text(cx + card_w / 2, card_y - 8 * mm, lbl, font_size=7, color=COLOR_GRAY, align='center')
            self._text(cx + card_w / 2, card_y - 18 * mm, val, font_size=15, color=COLOR_PRIMARY, align='center')
            self._text(cx + card_w / 2, card_y - 24 * mm, unit, font_size=7, color=COLOR_GRAY, align='center')

        # キーイベント合計バナー
        banner_y = card_y - card_h - 6 * mm
        banner_w = PAGE_W - 24 * mm
        self._rect(12 * mm, banner_y - 12 * mm, banner_w, 12 * mm, fill_color=COLOR_PRIMARY, radius=3)
        self._text(PAGE_W / 2, banner_y - 6 * mm,
                   f"月間キーイベント（コンバージョン）合計:  {kpis['key_events_total']:,}  件",
                   font_size=13, color=COLOR_WHITE, align='center')

        if daily_chart_bytes:
            chart_y = banner_y - 18 * mm
            chart_w = PAGE_W - 24 * mm
            self._image_bytes(daily_chart_bytes, 12 * mm, chart_y, chart_w)

        self._footer(company_name, 2)

    # ─── P3: 流入チャネル分析 ───
    def draw_channels(self, meta, kpis, bar_bytes, donut_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('流入チャネル分析')

        col_w = (PAGE_W - 36 * mm) / 2
        col1_x = 12 * mm
        col2_x = col1_x + col_w + 12 * mm
        content_top = PAGE_H - 30 * mm
        chart_h = 60 * mm

        self._section_header(col1_x, content_top, 'セッション数', 'チャネル別')
        if bar_bytes:
            self._image_bytes(bar_bytes, col1_x, content_top - 5 * mm, col_w, h=chart_h)

        self._section_header(col2_x, content_top, '新規ユーザー獲得元', 'チャネル別')
        if donut_bytes:
            self._image_bytes(donut_bytes, col2_x, content_top - 5 * mm, col_w, h=chart_h)

        # 下段: エンゲージメント品質指標カード
        breakdown_top = content_top - chart_h - 18 * mm
        self._section_header(col1_x, breakdown_top, 'エンゲージメント品質', 'サイト利用状況')

        card_data = [
            ('PV/セッション', f"{kpis.get('pv_per_session', 0):.1f}", 'ページ/回'),
            ('スクロール率', f"{kpis.get('scroll_rate', 0):.1f}%", ''),
            ('クリック率', f"{kpis.get('click_rate', 0):.1f}%", ''),
            ('ファイルDL', f"{kpis.get('file_download', 0):,}", '件'),
        ]
        num = len(card_data)
        total_w = PAGE_W - 24 * mm
        gap = 4 * mm
        cw = (total_w - (num - 1) * gap) / num
        ch = 25 * mm
        cy = breakdown_top - 15 * mm

        for i, (lbl, val, unit) in enumerate(card_data):
            cx = col1_x + i * (cw + gap)
            self.c.setFillColor(COLOR_CARD_BG)
            self.c.setStrokeColor(colors.HexColor('#DDE6EF'))
            self.c.roundRect(cx, cy - ch, cw, ch, 4, fill=1, stroke=1)
            self._rect(cx, cy - 3 * mm, cw, 3 * mm, fill_color=COLOR_SECONDARY, radius=0)
            self._text(cx + cw / 2, cy - 9 * mm, lbl, font_size=8, color=COLOR_GRAY, align='center')
            self._text(cx + cw / 2, cy - 18 * mm, val, font_size=16, color=COLOR_PRIMARY, align='center')
            self._text(cx + cw / 2, cy - 23 * mm, unit, font_size=7, color=COLOR_GRAY, align='center')

        # 海外アクセス注記
        non_jp = kpis.get('non_jp_users', 0)
        if non_jp > 0:
            note_y = cy - ch - 8 * mm
            self._text(col1_x, note_y,
                       f"※ 海外からのアクセス {non_jp:,}人を検出（ボットトラフィックの可能性あり。KPI値に含まれています）",
                       font_size=8, color=COLOR_GRAY)

        self._footer(company_name, 3)

    # ─── P4: キーイベント ───
    def draw_key_events(self, meta, kpis, kev_chart_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('キーイベント（コンバージョン）')
        content_top = PAGE_H - 30 * mm

        self._section_header(10 * mm, content_top, 'コンバージョン内訳')
        if kev_chart_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(kev_chart_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=75 * mm)

        kev = kpis.get('key_events', {})
        sessions = kpis.get('sessions', 0)

        # 成果CV（直接的な予約・問い合わせ）と補助CV（電話タップ等）を分離
        seika_map = {
            'WEB予約': 'WEB予約完了', 'web予約': 'WEB予約完了',
            '問診票_成人': '問診票送信（成人）', '問診票_小児': '問診票送信（小児）',
        }
        hojo_map = {
            'スマホ_TEL': '電話タップ',
            '成人電話': '電話タップ（成人）', '小児電話': '電話タップ（小児）',
        }

        def build_rows(event_map, label):
            rows = [[label, '件数', 'CVR']]
            total = 0
            for k, display in event_map.items():
                if k in kev and kev[k] > 0:
                    count = kev[k]
                    cvr = count / sessions * 100 if sessions > 0 else 0
                    rows.append([display, f"{count:,}", f"{cvr:.1f}%"])
                    total += count
            total_cvr = total / sessions * 100 if sessions > 0 else 0
            rows.append([f'{label} 小計', f"{total:,}", f"{total_cvr:.1f}%"])
            return rows, total

        seika_rows, seika_total = build_rows(seika_map, '成果CV')
        hojo_rows, hojo_total = build_rows(hojo_map, '補助CV')

        col_ws = [80 * mm, 25 * mm, 25 * mm]
        row_h = 7 * mm
        table_x = 10 * mm

        def draw_cv_table(rows, start_y):
            for ri, row in enumerate(rows):
                row_y = start_y - ri * row_h
                is_header = ri == 0
                is_total = ri == len(rows) - 1
                bg = COLOR_PRIMARY if is_header else \
                     colors.HexColor('#EEF3F8') if is_total else \
                     (COLOR_CARD_BG if ri % 2 == 0 else colors.HexColor('#F5F8FC'))
                self.c.setFillColor(bg)
                self.c.rect(table_x, row_y - row_h, sum(col_ws), row_h, fill=1, stroke=0)
                txt_color = COLOR_WHITE if is_header else COLOR_DARK
                for ci, (cell, cw) in enumerate(zip(row, col_ws)):
                    cell_x = table_x + sum(col_ws[:ci])
                    align = 'right' if ci >= 1 and not is_header else 'left'
                    pad = 3 * mm if align == 'left' else -3 * mm
                    self._text(cell_x + (cw + pad if align == 'right' else pad),
                               row_y - row_h + 2 * mm, str(cell),
                               font_size=9, color=txt_color, align=align)
            return start_y - len(rows) * row_h

        table_top = content_top - 95 * mm
        next_y = draw_cv_table(seika_rows, table_top)
        draw_cv_table(hojo_rows, next_y - 3 * mm)

        # 右側に成果CV/補助CVの説明
        note_x = table_x + sum(col_ws) + 10 * mm
        note_y = table_top - 5 * mm
        self._text(note_x, note_y, '成果CV', font_size=10, color=COLOR_PRIMARY)
        self._text(note_x, note_y - 7 * mm, '予約・問い合わせの直接的な成果', font_size=8, color=COLOR_DARK)
        self._text(note_x, note_y - 22 * mm, '補助CV', font_size=10, color=COLOR_ACCENT)
        self._text(note_x, note_y - 29 * mm, '電話タップ等の間接指標', font_size=8, color=COLOR_DARK)
        self._text(note_x, note_y - 36 * mm, '※電話タップ＝実際の通話・予約', font_size=7, color=COLOR_GRAY)
        self._text(note_x, note_y - 41 * mm, '  とは限りません', font_size=7, color=COLOR_GRAY)

        self._footer(company_name, 4)

    # ─── P5: 上位ページ ───
    def draw_pages(self, meta, pages_chart_bytes, df_pages, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('上位ページ（表示回数）')
        content_top = PAGE_H - 30 * mm
        self._section_header(10 * mm, content_top, '上位20ページ')

        if pages_chart_bytes:
            chart_w = 175 * mm
            self._image_bytes(pages_chart_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=155 * mm)

        if df_pages is not None and len(df_pages) > 0:
            col_p = df_pages.columns[0]
            col_v = df_pages.columns[1]
            df_top20 = df_pages.head(20).copy()

            def shorten(title, max_len=16):
                title = str(title)
                m = re.search(r'[「『](.*?)[」』]', title)
                if m:
                    s = m.group(1)
                    return s[:max_len] + '…' if len(s) > max_len else s
                parts = [p.strip() for p in re.split(r'[｜|]', title) if p.strip()]
                noise = r'(歯医者|歯科クリニック|歯科医院|武蔵新城|新丸子|川崎市|北本市)'
                meaningful = [p for p in parts if not re.search(noise, p)]
                if meaningful:
                    s = meaningful[0]
                elif len(parts) >= 2:
                    s = 'トップページ'
                else:
                    s = parts[-1] if parts else title
                s = re.sub(r'^(川崎市で|北本市で)', '', s).strip()
                return s[:max_len] + '…' if len(s) > max_len else s

            rows = [['順位', 'ページ', '表示']]
            for i, (_, r) in enumerate(df_top20.iterrows(), start=1):
                rows.append([str(i), shorten(r[col_p]), f"{int(r[col_v]):,}"])

            tx = 198 * mm
            ty = content_top - 2 * mm
            rh = 6.7 * mm
            cws = [10 * mm, 48 * mm, 16 * mm]

            for ri, row in enumerate(rows):
                ry = ty - ri * rh
                is_header = ri == 0
                bg = COLOR_PRIMARY if is_header else (COLOR_CARD_BG if ri % 2 == 1 else colors.HexColor('#F5F8FC'))
                txt_color = COLOR_WHITE if is_header else COLOR_DARK
                x = tx
                for ci, (cell, cw) in enumerate(zip(row, cws)):
                    self.c.setFillColor(bg)
                    self.c.setStrokeColor(colors.HexColor('#DDE6EF'))
                    self.c.rect(x, ry - rh, cw, rh, fill=1, stroke=1)
                    align = 'center' if ci == 0 else ('right' if ci == 2 else 'left')
                    if align == 'center':
                        self._text(x + cw / 2, ry - rh + 2 * mm, cell, font_size=8, color=txt_color, align='center')
                    elif align == 'right':
                        self._text(x + cw - 2 * mm, ry - rh + 2 * mm, cell, font_size=8, color=txt_color, align='right')
                    else:
                        self._text(x + 1.5 * mm, ry - rh + 2 * mm, cell, font_size=8, color=txt_color)
                    x += cw

        self._footer(company_name, 5)

    # ─── P6: CVR分析 + フォーム完了率 ───
    def draw_cvr_analysis(self, kpis, cvr_chart_bytes, form_funnel_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('CVR（コンバージョン率）分析')
        content_top = PAGE_H - 30 * mm

        self._section_header(10 * mm, content_top, 'コンバージョン率分析', 'セッション数ベース')
        if cvr_chart_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(cvr_chart_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=60 * mm)

        # フォーム完了率
        form_top = content_top - 75 * mm
        self._section_header(10 * mm, form_top, 'お問い合わせフォーム完了率')
        if form_funnel_bytes:
            self._image_bytes(form_funnel_bytes, 10 * mm, form_top - 5 * mm, PAGE_W / 2 - 20 * mm, h=55 * mm)

        # 右側にインサイトテキスト
        insight_x = PAGE_W / 2 + 10 * mm
        insight_y = form_top - 5 * mm
        self._text(insight_x, insight_y, 'ポイント', font_size=10, color=COLOR_PRIMARY)

        form_start = kpis.get('form_start', 0)
        form_submit = kpis.get('form_submit', 0)
        form_start_rate = kpis.get('form_start_to_submit', 0)

        insights = [
            f"・全体CVR {kpis['cvr_total']:.1f}%",
            f"  （キーイベント{kpis['key_events_total']:,}件 ÷ セッション{kpis['sessions']:,}）",
            f"・フォームページ閲覧→完了率 {kpis['form_completion_rate']:.1f}%",
            "",
            "【フォーム操作の離脱分析】",
            f"・フォーム入力開始: {form_start:,}回",
            f"・フォーム送信完了: {form_submit:,}回",
            f"・入力→送信完了率: {form_start_rate:.1f}%",
        ]
        if form_start > 0 and form_start_rate < 10:
            insights.append("")
            insights.append("⚠ フォーム入力後の離脱が多い。")
            insights.append("  入力項目数・エラー表示・送信ボタンの")
            insights.append("  視認性を確認推奨。")

        line_h = 6.5 * mm
        for i, text in enumerate(insights):
            clr = COLOR_RED if text.startswith('⚠') else COLOR_DARK
            self._text(insight_x, insight_y - (i + 1) * line_h, text, font_size=7.5, color=clr)

        self._footer(company_name, 6)

    # ─── P7: ユーザー分析（リテンション + 新規vs既存） ───
    def draw_user_analysis(self, kpis, retention_bytes, new_vs_returning_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('ユーザー分析')
        content_top = PAGE_H - 30 * mm

        # リテンション
        self._section_header(10 * mm, content_top, 'ユーザー維持率', '週別コホート分析')
        if retention_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(retention_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=65 * mm)

        # 新規 vs リピーター
        lower_top = content_top - 80 * mm
        col_w = (PAGE_W - 36 * mm) / 2
        col1_x = 12 * mm
        col2_x = col1_x + col_w + 12 * mm

        self._section_header(col1_x, lower_top, '新規 vs リピーター', 'ユーザー構成')
        if new_vs_returning_bytes:
            self._image_bytes(new_vs_returning_bytes, col1_x, lower_top - 5 * mm, col_w, h=65 * mm)

        # 右側: インサイト
        self._section_header(col2_x, lower_top, 'データサマリ')
        insight_y = lower_top - 15 * mm
        new_ratio = kpis.get('new_user_ratio', 0)
        returning_ratio = 100 - new_ratio
        returning_users = kpis.get('active_users', 0) - kpis.get('new_users', 0)
        insights = [
            f"新規ユーザー比率: {new_ratio:.0f}%",
            f"リピーター比率: {returning_ratio:.0f}%（{returning_users:,}人）",
            "",
            f"新規ユーザー: {kpis.get('new_users', 0):,}人",
            f"アクティブUU: {kpis.get('active_users', 0):,}人",
            "",
            "※新規比率が高い場合、初回訪問から",
            "  予約に至る導線の最適化が重要です。",
            "※リピーター比率はブランド認知や",
            "  再訪促進施策の効果指標になります。",
        ]
        for i, text in enumerate(insights):
            fs = 10 if i == 0 else 9
            clr = COLOR_PRIMARY if i == 0 else COLOR_DARK
            self._text(col2_x + 5 * mm, insight_y - i * 8 * mm, text, font_size=fs, color=clr)

        self._footer(company_name, 7)

    # ─── P8: コンテンツカテゴリ + 曜日分析 ───
    def draw_content_and_weekday(self, kpis, content_chart_bytes, weekday_chart_bytes, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('コンテンツ・曜日分析')
        content_top = PAGE_H - 30 * mm

        # コンテンツカテゴリ
        self._section_header(10 * mm, content_top, 'コンテンツカテゴリ分析', 'ページタイトルから自動分類')
        if content_chart_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(content_chart_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=70 * mm)

        # 曜日分析
        weekday_top = content_top - 85 * mm
        self._section_header(10 * mm, weekday_top, '曜日別アクセスパターン', '広告出稿・コンテンツ配信の最適化に')
        if weekday_chart_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(weekday_chart_bytes, 10 * mm, weekday_top - 5 * mm, chart_w, h=65 * mm)

        self._footer(company_name, 8)

    # ─── P9: Paid Social分析 ───
    def draw_paid_social(self, kpis, paid_social_chart_bytes, df_ch_ses, df_ch_new, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('Paid Social（有料SNS広告）効果分析')
        content_top = PAGE_H - 30 * mm

        self._section_header(10 * mm, content_top, 'Paid Social 効果')
        if paid_social_chart_bytes:
            chart_w = PAGE_W - 20 * mm
            self._image_bytes(paid_social_chart_bytes, 10 * mm, content_top - 5 * mm, chart_w, h=70 * mm)

        # KPIカード
        card_top = content_top - 85 * mm
        self._section_header(10 * mm, card_top, '広告効果サマリ')

        # Paid Socialのデータ抽出
        ps_sessions = 0
        ps_new_users = 0
        total_sessions = kpis.get('sessions', 0)
        if df_ch_ses is not None:
            col_ch = df_ch_ses.columns[0]
            col_val = df_ch_ses.columns[1]
            ps_row = df_ch_ses[df_ch_ses[col_ch] == 'Paid Social']
            if not ps_row.empty:
                ps_sessions = int(ps_row[col_val].iloc[0])
        if df_ch_new is not None:
            col_ch = df_ch_new.columns[0]
            col_val = df_ch_new.columns[1]
            ps_row = df_ch_new[df_ch_new[col_ch] == 'Paid Social']
            if not ps_row.empty:
                ps_new_users = int(ps_row[col_val].iloc[0])

        ps_session_ratio = ps_sessions / total_sessions * 100 if total_sessions > 0 else 0

        cards = [
            ('Paid Socialセッション', f'{ps_sessions:,}', '件'),
            ('セッション構成比', f'{ps_session_ratio:.1f}%', ''),
            ('Paid Social新規ユーザー', f'{ps_new_users:,}', '人'),
        ]

        margin = 10 * mm
        card_w = (PAGE_W - 2 * margin - 2 * 4 * mm) / 3
        card_h = 25 * mm
        for i, (lbl, val, unit) in enumerate(cards):
            cx = margin + i * (card_w + 4 * mm)
            cy = card_top - 10 * mm
            self.c.setFillColor(COLOR_CARD_BG)
            self.c.setStrokeColor(colors.HexColor('#DDE6EF'))
            self.c.roundRect(cx, cy - card_h, card_w, card_h, 4, fill=1, stroke=1)
            self._rect(cx, cy - 3 * mm, card_w, 3 * mm, fill_color=COLOR_ACCENT, radius=0)
            self._text(cx + card_w / 2, cy - 9 * mm, lbl, font_size=8, color=COLOR_GRAY, align='center')
            self._text(cx + card_w / 2, cy - 18 * mm, val, font_size=16, color=COLOR_ACCENT, align='center')
            self._text(cx + card_w / 2, cy - 23 * mm, unit, font_size=8, color=COLOR_GRAY, align='center')

        self._footer(company_name, 9)

    # ─── P10: 計測定義・注記 ───
    def draw_notes(self, meta, kpis, company_name=''):
        self._bg(COLOR_LIGHT)
        self._page_title_bar('計測定義・注記')
        content_top = PAGE_H - 30 * mm
        lx = 12 * mm
        line_h = 6 * mm
        note_line_h = 6.5 * mm

        # KPI定義
        self._section_header(lx, content_top, 'KPI定義')
        y = content_top - 10 * mm
        defs = [
            ['指標', '定義'],
            ['月間アクティブユーザー', 'GA4「30日アクティブユーザー」の期間最終日の値'],
            ['新規ユーザー', '期間内の日別新規ユーザー数の合計'],
            ['セッション数', '全チャネルのセッション数合計'],
            ['全体CVR', 'キーイベント合計 ÷ セッション数 × 100（%）'],
            ['成果CV', 'WEB予約完了・お問い合わせフォーム送信完了'],
            ['補助CV', '電話タップ（実際の通話・予約成立とは異なる）'],
            ['フォーム完了率', 'お問い合わせ完了PV ÷ お問い合わせページPV × 100（%）'],
            ['新規ユーザー比率', '新規ユーザー数 ÷ アクティブユーザー数 × 100（%）'],
            ['PV/セッション', '総ページビュー数 ÷ セッション数'],
            ['スクロール率', 'scrollイベント数 ÷ 総PV数 × 100（%）'],
        ]
        col_ws = [55 * mm, PAGE_W - 24 * mm - 55 * mm]
        for ri, row in enumerate(defs):
            ry = y - ri * line_h
            is_header = ri == 0
            bg = COLOR_PRIMARY if is_header else (COLOR_CARD_BG if ri % 2 == 1 else colors.HexColor('#F5F8FC'))
            txt_color = COLOR_WHITE if is_header else COLOR_DARK
            self.c.setFillColor(bg)
            self.c.rect(lx, ry - line_h, sum(col_ws), line_h, fill=1, stroke=0)
            self._text(lx + 3 * mm, ry - line_h + 1.5 * mm, row[0], font_size=7.5, color=txt_color)
            self._text(lx + col_ws[0] + 3 * mm, ry - line_h + 1.5 * mm, row[1], font_size=7.5, color=txt_color)

        # キーイベント計測条件
        note_y = y - len(defs) * line_h - 12 * mm
        self._section_header(lx, note_y, 'キーイベント計測条件')
        notes = [
            f"・WEB予約: GA4イベント「web予約」（予約フォーム完了時にトリガー）",
            f"・電話タップ: GA4イベント「スマホ_TEL」（電話番号リンクのタップ時にトリガー。実際の発信・通話を保証するものではありません）",
            f"・フォーム開始/送信: GA4自動収集イベント「form_start」「form_submit」",
        ]
        for i, note in enumerate(notes):
            self._text(lx + 5 * mm, note_y - (i + 1) * note_line_h, note, font_size=7.5, color=COLOR_DARK)

        # データ取得条件
        cond_y = note_y - (len(notes) + 1) * note_line_h - 6 * mm
        self._section_header(lx, cond_y, 'データ取得条件')
        period = ''
        if meta.get('start') and meta.get('end'):
            s, e = meta['start'], meta['end']
            period = f"{s[:4]}/{s[4:6]}/{s[6:]} ～ {e[:4]}/{e[4:6]}/{e[6:]}"
        conditions = [
            f"・対象期間: {period}",
            f"・データソース: Google Analytics 4 標準エクスポートCSV",
            f"・プロパティ: {meta.get('property', '')}",
            f"・内部トラフィック除外: GA4側の設定に依存",
            f"・海外アクセス: 除外せず集計に含む（{kpis.get('non_jp_users', 0):,}人を検出）",
            f"・チャネル別CV数: GA4標準CSVにはチャネル×CVのクロスデータが含まれないため全体値のみ表示",
        ]
        for i, cond in enumerate(conditions):
            self._text(lx + 5 * mm, cond_y - (i + 1) * note_line_h, cond, font_size=7.5, color=COLOR_DARK)

        self._footer(company_name, 10)


# ─────────────────────────────────────────────────────
# メイン処理
# ─────────────────────────────────────────────────────
def generate_report(csv_path: str, output_path: str,
                    company_name: str = 'YOUR COMPANY',
                    department: str  = '',
                    logo_path: str   = None) -> str:
    """GA4 CSVからPDFレポート（高度解析版）を生成"""
    print(f"[1/8] CSVを読み込み中: {csv_path}")
    data = parse_ga4_csv(csv_path)
    meta = data['meta']
    kpis = data['kpis']

    print(f"[2/8] KPI集計完了 | 期間: {meta['month_label']}")
    print(f"      アクティブUU: {kpis['active_users']:,}  新規UU: {kpis['new_users']:,}"
          f"  セッション: {kpis['sessions']:,}  CVR: {kpis['cvr_total']:.1f}%"
          f"  キーイベント: {kpis['key_events_total']:,}")

    print("[3/8] 標準チャートを生成中...")
    daily_chart   = make_daily_line_chart(data['df_active'], data['df_new'], meta['month_label'])
    bar_chart     = make_channel_bar_chart(data['df_ch_ses'])
    donut_chart   = make_channel_donut_chart(data['df_ch_new'])
    kev_chart     = make_key_events_bar(kpis['key_events'])
    pages_chart   = make_top_pages_chart(data['df_pages'], top_n=20)

    print("[4/8] 高度解析チャートを生成中...")
    cvr_chart = make_cvr_gauge_chart(kpis['cvr_total'], kpis['cvr_detail'], kpis['sessions'])
    retention_chart = make_retention_heatmap(data['df_retention'])
    content_chart = make_content_category_chart(kpis['content_categories'])
    weekday_chart = make_weekday_chart(data['df_active'], data['df_new'], meta.get('start', ''))
    new_vs_returning_chart = make_new_vs_returning_chart(kpis['new_users'], kpis['active_users'])
    paid_social_chart = make_paid_social_chart(data['df_ch_ses'], data['df_ch_new'])
    form_funnel_chart = None
    if kpis.get('contact_pv', 0) > 0:
        form_funnel_chart = make_form_funnel_chart(
            kpis['contact_pv'], kpis['contact_done_pv'], kpis['form_completion_rate'])

    print("[5/9] PDFを生成中...")
    rc = ReportCanvas(output_path)

    # P1: 表紙
    rc.draw_cover(meta, company_name, department, logo_path)
    rc.new_page()

    # P2: サマリ
    rc.draw_summary(meta, kpis, daily_chart, company_name)
    rc.new_page()

    # P3: チャネル分析
    rc.draw_channels(meta, kpis, bar_chart, donut_chart, company_name)
    rc.new_page()

    # P4: キーイベント
    rc.draw_key_events(meta, kpis, kev_chart, company_name)
    rc.new_page()

    # P5: 上位ページ
    rc.draw_pages(meta, pages_chart, data['df_pages'], company_name)
    rc.new_page()

    # P6: CVR分析
    rc.draw_cvr_analysis(kpis, cvr_chart, form_funnel_chart, company_name)
    rc.new_page()

    # P7: ユーザー分析（リテンション）
    rc.draw_user_analysis(kpis, retention_chart, new_vs_returning_chart, company_name)
    rc.new_page()

    # P8: コンテンツ・曜日分析
    rc.draw_content_and_weekday(kpis, content_chart, weekday_chart, company_name)
    rc.new_page()

    # P9: Paid Social分析
    rc.draw_paid_social(kpis, paid_social_chart, data['df_ch_ses'], data['df_ch_new'], company_name)
    rc.new_page()

    # P10: 計測定義・注記
    rc.draw_notes(meta, kpis, company_name)

    rc.save()
    print(f"[6/9] PDF保存完了: {output_path}")
    return output_path


# ─────────────────────────────────────────────────────
# CLI エントリポイント
# ─────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使い方: python3 report_logic_advanced.py <CSVファイル> [出力PDF] [会社名] [部署名] [ロゴパス]")
        sys.exit(1)

    csv_file    = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else csv_file.replace('.csv', '_advanced_report.pdf')
    company     = sys.argv[3] if len(sys.argv) > 3 else 'YOUR COMPANY'
    dept        = sys.argv[4] if len(sys.argv) > 4 else ''
    logo        = sys.argv[5] if len(sys.argv) > 5 else None

    result = generate_report(csv_file, output_file, company, dept, logo)
    print(f"[完了] {result}")
