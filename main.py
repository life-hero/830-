# -*- coding: utf-8 -*-

"""
分组统计工具 – 图形界面版本（基于 tkinter）
支持从原始数据生成统计报表（Excel + HTML），或从已生成的 Excel 重建 HTML。
"""

import webbrowser
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from collections import defaultdict
from typing import List, Dict, Any, Optional, Tuple
import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import time

# ---------- 常量 ----------
FONT_SIZE_HTML = 12
DEBUG = True  # 开启调试打印


def debug_print(msg: str) -> None:
    """调试打印函数"""
    if DEBUG:
        timestamp = time.strftime("%H:%M:%S")
        print(f"[{timestamp}] {msg}")


# ---------- 分类映射 ----------
def classify_main_type(val: str) -> str:
    """
    将产品主推类型映射为三大分类
    TOP+/TOP是标准化，NON-TOP是配置化
    """
    if val in ['TOP', 'TOP+']:
        return '标准化'
    elif val == 'NON-TOP':
        return '配置化'
    elif val == '定制':
        return '定制化'
    else:
        return '其他'


# ---------- 按销售大区统计 ----------
def compute_region_stats(df: pd.DataFrame) -> Tuple[List[str], Dict[str, Any]]:
    debug_print("开始 compute_region_stats")
    df_temp = df.copy()
    df_temp['主推分类'] = df_temp['产品主推类型（产品维度）'].apply(classify_main_type)

    region_class = df_temp.groupby(['销售大区', '主推分类'])['总面积'].sum().unstack(fill_value=0)
    for cat in ['标准化', '配置化', '定制化']:
        if cat not in region_class.columns:
            region_class[cat] = 0
    region_class['面积汇总'] = region_class.sum(axis=1)

    total_area = region_class['面积汇总'].sum()
    region_class['占比'] = (region_class['面积汇总'] / total_area * 100) if total_area else 0
    region_class_sorted = region_class.sort_values('面积汇总', ascending=False)
    region_class_sorted['排名'] = range(1, len(region_class_sorted) + 1)

    sorted_regions = region_class_sorted.index.tolist()
    region_data = {}
    for region in sorted_regions:
        row = region_class_sorted.loc[region]
        region_data[region] = {
            '面积汇总': row['面积汇总'],
            '标准化': row['标准化'],
            '配置化': row['配置化'],
            '定制化': row['定制化'],
            '占比': row['占比'],
            '排名': row['排名']
        }
    debug_print("compute_region_stats 完成")
    return sorted_regions, region_data


# ---------- 核心数据处理 ----------
def process_region_data_bkkk(df: pd.DataFrame) -> pd.DataFrame:
    debug_print("开始 process_region_data")
    # 1. 按平台类别+产品系列+产品型号汇总面积和订单数
    grouped = df.groupby(['平台类别', '产品系列', '产品型号'], as_index=False).agg({
        '总面积': 'sum',
        '订单数量': 'sum'
    })

    # 2. 合并主推分类信息，并添加定制原因
    type_map = df[['产品型号', '产品主推类型（产品维度）', '主推分类']].drop_duplicates(subset='产品型号')
    grouped = grouped.merge(type_map, on='产品型号', how='left')
    grouped['定制原因'] = grouped['产品主推类型（产品维度）'].apply(lambda x: '定制' if x == '定制' else '')

    # 3. 透视成三列（标准化/配置化/定制化）
    pivot_df = grouped.pivot_table(
        index=['平台类别', '产品系列', '产品型号'],
        columns='主推分类',
        values='总面积',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    for cat in ['标准化', '配置化', '定制化']:
        if cat not in pivot_df.columns:
            pivot_df[cat] = 0

    pivot_df['面积汇总'] = pivot_df[['标准化', '配置化', '定制化']].sum(axis=1)

    custom_map = grouped[['平台类别', '产品系列', '产品型号', '定制原因']].drop_duplicates()
    pivot_df = pivot_df.merge(custom_map, on=['平台类别', '产品系列', '产品型号'], how='left')

    # --- 按销售大区统计明细 ---
    region_stats = compute_region_stats(df)
    sorted_regions, region_data = region_stats

    region_group = df.groupby(['平台类别', '产品系列', '产品型号', '销售大区', '主推分类']).agg({
        '总面积': 'sum'
    }).reset_index()

    region_pivot = region_group.pivot_table(
        index=['平台类别', '产品系列', '产品型号'],
        columns=['销售大区', '主推分类'],
        values='总面积',
        fill_value=0
    ).reset_index()

    key_cols = ['平台类别', '产品系列', '产品型号']
    keys = pd.DataFrame()
    for col in key_cols:
        found = False
        for c in region_pivot.columns:
            if col in str(c):
                keys[col] = region_pivot[c]
                found = True
                break
        if not found:
            raise KeyError(f"在 region_pivot 中找不到列: {col}")

    stat_cols = region_pivot.drop(columns=list(keys.columns))
    stat_cols.columns = [
        f"{col[0]}_{col[1]}" if isinstance(col, tuple) else str(col)
        for col in stat_cols.columns
    ]
    region_pivot_flat = pd.concat([keys, stat_cols], axis=1)

    pivot_df = pivot_df.merge(region_pivot_flat, on=key_cols, how='left')

    region_cols = []
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            region_cols.append(f"{region}_{cat}")

    for col in region_cols:
        if col not in pivot_df.columns:
            pivot_df[col] = 0
        else:
            pivot_df[col] = pivot_df[col].fillna(0)

    pivot_df['总计'] = pivot_df[region_cols].sum(axis=1)

    base_cols = ['平台类别', '产品系列', '产品型号', '面积汇总', '标准化', '配置化', '定制化', '定制原因']
    final_cols = base_cols + region_cols + ['总计']
    for col in final_cols:
        if col not in pivot_df.columns:
            pivot_df[col] = 0

    final_df = pivot_df[final_cols]

    numeric_cols = ['面积汇总', '标准化', '配置化', '定制化'] + region_cols + ['总计']
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = final_df[col].round(0).astype(int)

    final_df['面积汇总'] = final_df['标准化'] + final_df['配置化'] + final_df['定制化']

    platform_totals = final_df.groupby('平台类别')['面积汇总'].sum().to_dict()
    final_df['_platform_total'] = final_df['平台类别'].map(platform_totals)

    series_totals = final_df.groupby(['平台类别', '产品系列'])['面积汇总'].sum().to_dict()
    final_df['_series_total'] = final_df.apply(
        lambda r: series_totals.get((r['平台类别'], r['产品系列']), 0), axis=1
    )

    final_df = final_df.sort_values(
        ['_platform_total', '_series_total', '面积汇总'],
        ascending=[False, False, False]
    ).drop(columns=['_platform_total', '_series_total']).reset_index(drop=True)

    total_area = final_df['面积汇总'].sum()
    total_std = final_df['标准化'].sum()
    total_cfg = final_df['配置化'].sum()
    total_cus = final_df['定制化'].sum()

    std_ratio = (total_std / total_area * 100) if total_area else 0
    cfg_ratio = (total_cfg / total_area * 100) if total_area else 0
    cus_ratio = (total_cus / total_area * 100) if total_area else 0
    col_std = f"标准化({std_ratio:.1f}%)"
    col_cfg = f"配置化({cfg_ratio:.1f}%)"
    col_cus = f"定制化({cus_ratio:.1f}%)"

    rows = []

    total_row_dict = {
        '平台类别': '总计',
        '产品系列': '',
        '产品型号': '',
        '面积汇总': total_area,
        '标准化': total_std,
        '配置化': total_cfg,
        '定制化': total_cus,
        '定制原因': ''
    }
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            col_name = f"{region}_{cat}"
            total_row_dict[col_name] = final_df[col_name].sum()
    total_row_dict['总计'] = final_df['总计'].sum()

    top_total_row = total_row_dict.copy()
    top_total_row['平台类别'] = ''
    top_total_row['产品系列'] = ''
    top_total_row['产品型号'] = ''
    rows.append(top_total_row)

    for plat, group in final_df.groupby('平台类别', sort=False):
        for _, row in group.iterrows():
            rows.append(row.to_dict())
        g_sum_area = group['面积汇总'].sum()
        g_sum_std = group['标准化'].sum()
        g_sum_cfg = group['配置化'].sum()
        g_sum_cus = group['定制化'].sum()
        ratio = (g_sum_area / total_area * 100) if total_area != 0 else 0
        ratio_text = f"占比 {ratio:.2f}%"
        subtotal = {
            '平台类别': f"{plat}汇总",
            '产品系列': ratio_text,
            '产品型号': '',
            '面积汇总': g_sum_area,
            '标准化': g_sum_std,
            '配置化': g_sum_cfg,
            '定制化': g_sum_cus,
            '定制原因': ''
        }
        for region in sorted_regions:
            for cat in ['标准化', '配置化', '定制化']:
                col_name = f"{region}_{cat}"
                if col_name in group.columns:
                    subtotal[col_name] = group[col_name].sum()
                else:
                    subtotal[col_name] = 0
        subtotal['总计'] = sum(subtotal[col] for col in region_cols)
        rows.append(subtotal)

    rows.append(total_row_dict)

    result_df = pd.DataFrame(rows)

    new_cols = ['平台类别', '产品系列', '产品型号', '面积汇总', col_std, col_cfg, col_cus, '定制原因']
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            new_cols.append(f"{region}_{cat}")
    new_cols.append('总计')
    result_df.columns = new_cols

    result_df.attrs['sorted_regions'] = sorted_regions
    result_df.attrs['region_data'] = region_data
    result_df.attrs['region_area'] = {r: region_data[r]['面积汇总'] for r in sorted_regions}
    result_df.attrs['region_ratio'] = {r: region_data[r]['占比'] for r in sorted_regions}
    result_df.attrs['region_rank'] = {r: region_data[r]['排名'] for r in sorted_regions}

    debug_print("process_region_data 完成")
    return result_df


def process_region_data(df: pd.DataFrame) -> pd.DataFrame:
    debug_print("开始 process_region_data")

    # 1. 按 平台类别+产品系列+产品型号+主推分类 分组汇总面积
    grouped = df.groupby(
        ['平台类别', '产品系列', '产品型号', '主推分类'],
        as_index=False
    ).agg({
        '总面积': 'sum'
        # 订单数量如需可保留，但后续未使用，这里暂不处理
    })

    # 2. 透视，将主推分类展开为列
    pivot_df = grouped.pivot_table(
        index=['平台类别', '产品系列', '产品型号'],
        columns='主推分类',
        values='总面积',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # 确保三列存在（标准化、配置化、定制化）
    for cat in ['标准化', '配置化', '定制化']:
        if cat not in pivot_df.columns:
            pivot_df[cat] = 0

    # 3. 提取定制原因（若某型号存在主推类型为“定制”的记录）
    custom_df = df[df['产品主推类型（产品维度）'] == '定制'][['产品型号']].drop_duplicates()
    custom_df['定制原因'] = '定制'
    pivot_df = pivot_df.merge(custom_df, on='产品型号', how='left')
    pivot_df['定制原因'] = pivot_df['定制原因'].fillna('')

    # 4. 计算面积汇总（三列之和）
    pivot_df['面积汇总'] = pivot_df[['标准化', '配置化', '定制化']].sum(axis=1)

    # ---- 以下区域统计明细保持不变（region_group 本身已按主推分类分组） ----
    region_stats = compute_region_stats(df)
    sorted_regions, region_data = region_stats

    region_group = df.groupby(['平台类别', '产品系列', '产品型号', '销售大区', '主推分类']).agg({
        '总面积': 'sum'
    }).reset_index()

    region_pivot = region_group.pivot_table(
        index=['平台类别', '产品系列', '产品型号'],
        columns=['销售大区', '主推分类'],
        values='总面积',
        fill_value=0
    ).reset_index()

    # 提取键列（平台类别、产品系列、产品型号）与统计列合并
    key_cols = ['平台类别', '产品系列', '产品型号']
    keys = pd.DataFrame()
    for col in key_cols:
        found = False
        for c in region_pivot.columns:
            if col in str(c):
                keys[col] = region_pivot[c]
                found = True
                break
        if not found:
            raise KeyError(f"在 region_pivot 中找不到列: {col}")

    stat_cols = region_pivot.drop(columns=list(keys.columns))
    stat_cols.columns = [
        f"{col[0]}_{col[1]}" if isinstance(col, tuple) else str(col)
        for col in stat_cols.columns
    ]
    region_pivot_flat = pd.concat([keys, stat_cols], axis=1)

    # 合并到主透视表
    pivot_df = pivot_df.merge(region_pivot_flat, on=key_cols, how='left')

    # 补齐区域列（确保所有区域都存在）
    region_cols = []
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            region_cols.append(f"{region}_{cat}")

    for col in region_cols:
        if col not in pivot_df.columns:
            pivot_df[col] = 0
        else:
            pivot_df[col] = pivot_df[col].fillna(0)

    # 计算总计（区域所有分类之和）
    pivot_df['总计'] = pivot_df[region_cols].sum(axis=1)

    # 构建最终列顺序
    base_cols = ['平台类别', '产品系列', '产品型号', '面积汇总', '标准化', '配置化', '定制化', '定制原因']
    final_cols = base_cols + region_cols + ['总计']
    for col in final_cols:
        if col not in pivot_df.columns:
            pivot_df[col] = 0

    final_df = pivot_df[final_cols]

    # 数值列取整
    numeric_cols = ['面积汇总', '标准化', '配置化', '定制化'] + region_cols + ['总计']
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = final_df[col].round(0).astype(int)

    # 重新计算面积汇总（确保一致）
    final_df['面积汇总'] = final_df['标准化'] + final_df['配置化'] + final_df['定制化']

    platform_totals = final_df.groupby('平台类别')['面积汇总'].sum().to_dict()
    final_df['_platform_total'] = final_df['平台类别'].map(platform_totals)

    series_totals = final_df.groupby(['平台类别', '产品系列'])['面积汇总'].sum().to_dict()
    final_df['_series_total'] = final_df.apply(
        lambda r: series_totals.get((r['平台类别'], r['产品系列']), 0), axis=1
    )

    final_df = final_df.sort_values(
        ['_platform_total', '_series_total', '面积汇总'],
        ascending=[False, False, False]
    ).drop(columns=['_platform_total', '_series_total']).reset_index(drop=True)

    total_area = final_df['面积汇总'].sum()
    total_std = final_df['标准化'].sum()
    total_cfg = final_df['配置化'].sum()
    total_cus = final_df['定制化'].sum()

    std_ratio = (total_std / total_area * 100) if total_area else 0
    cfg_ratio = (total_cfg / total_area * 100) if total_area else 0
    cus_ratio = (total_cus / total_area * 100) if total_area else 0
    col_std = f"标准化({std_ratio:.1f}%)"
    col_cfg = f"配置化({cfg_ratio:.1f}%)"
    col_cus = f"定制化({cus_ratio:.1f}%)"

    rows = []

    total_row_dict = {
        '平台类别': '总计',
        '产品系列': '',
        '产品型号': '',
        '面积汇总': total_area,
        '标准化': total_std,
        '配置化': total_cfg,
        '定制化': total_cus,
        '定制原因': ''
    }
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            col_name = f"{region}_{cat}"
            total_row_dict[col_name] = final_df[col_name].sum()
    total_row_dict['总计'] = final_df['总计'].sum()

    top_total_row = total_row_dict.copy()
    top_total_row['平台类别'] = ''
    top_total_row['产品系列'] = ''
    top_total_row['产品型号'] = ''
    rows.append(top_total_row)

    for plat, group in final_df.groupby('平台类别', sort=False):
        for _, row in group.iterrows():
            rows.append(row.to_dict())
        g_sum_area = group['面积汇总'].sum()
        g_sum_std = group['标准化'].sum()
        g_sum_cfg = group['配置化'].sum()
        g_sum_cus = group['定制化'].sum()
        ratio = (g_sum_area / total_area * 100) if total_area != 0 else 0
        ratio_text = f"占比 {ratio:.2f}%"
        subtotal = {
            '平台类别': f"{plat}汇总",
            '产品系列': ratio_text,
            '产品型号': '',
            '面积汇总': g_sum_area,
            '标准化': g_sum_std,
            '配置化': g_sum_cfg,
            '定制化': g_sum_cus,
            '定制原因': ''
        }
        for region in sorted_regions:
            for cat in ['标准化', '配置化', '定制化']:
                col_name = f"{region}_{cat}"
                if col_name in group.columns:
                    subtotal[col_name] = group[col_name].sum()
                else:
                    subtotal[col_name] = 0
        subtotal['总计'] = sum(subtotal[col] for col in region_cols)
        rows.append(subtotal)

    rows.append(total_row_dict)

    result_df = pd.DataFrame(rows)

    new_cols = ['平台类别', '产品系列', '产品型号', '面积汇总', col_std, col_cfg, col_cus,
                '定制原因']
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            new_cols.append(f"{region}_{cat}")
    new_cols.append('总计')
    result_df.columns = new_cols

    result_df.attrs['sorted_regions'] = sorted_regions
    result_df.attrs['region_data'] = region_data
    result_df.attrs['region_area'] = {r: region_data[r]['面积汇总'] for r in sorted_regions}
    result_df.attrs['region_ratio'] = {r: region_data[r]['占比'] for r in sorted_regions}
    result_df.attrs['region_rank'] = {r: region_data[r]['排名'] for r in sorted_regions}

    debug_print("process_region_data 完成")

    return result_df


# ---------- Excel 合并与样式 ----------
def apply_merges(
    sheet,
    start_row: int = 3,
    sorted_regions: List[str] = None,
    region_data: Dict[str, Any] = None,
    col_names: List[str] = None
) -> None:
    debug_print("开始 apply_merges")
    max_row = sheet.max_row
    max_col = sheet.max_column
    if max_row < start_row:
        return

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    if col_names is None:
        col_names = [sheet.cell(row=1, column=i).value for i in range(1, 9)]
        if not any(col_names):
            col_names = ['平台类别', '产品系列', '产品型号', '面积汇总', '标准化', '配置化', '定制化', '定制原因']

    for col_idx, name in enumerate(col_names[:8], start=1):
        sheet.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)
        cell = sheet.cell(row=1, column=col_idx)
        cell.value = name
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    if sorted_regions and region_data:
        start_col = 9
        total_col_index = start_col + len(sorted_regions) * 3

        for i, region in enumerate(sorted_regions):
            col_start = start_col + i * 3
            col_end = col_start + 2
            sheet.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
            info = region_data.get(region, {})
            rank = info.get('排名', 0)
            area = info.get('面积汇总', 0)
            ratio = info.get('占比', 0)
            display_text = f"TOP{rank:.0f} {region}\n{area:.2f}㎡ ({ratio:.1f}%)"
            cell = sheet.cell(row=1, column=col_start)
            cell.value = display_text
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border

            for j, sub in enumerate(['标准化', '配置化', '定制化']):
                cell = sheet.cell(row=2, column=col_start + j)
                cell.value = sub
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border

        sheet.merge_cells(start_row=1, start_column=total_col_index, end_row=2, end_column=total_col_index)
        cell = sheet.cell(row=1, column=total_col_index)
        cell.value = '总计'
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

    if max_row >= 3:
        sheet.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
        cell = sheet.cell(row=3, column=1)
        cell.value = '总计'
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        for col in range(4, max_col + 1):
            cell = sheet.cell(row=3, column=col)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            cell.font = Font(bold=True)

    detail_rows = []
    for r in range(start_row, max_row + 1):
        plat_val = sheet.cell(row=r, column=1).value
        if plat_val is None:
            continue
        if isinstance(plat_val, str) and (plat_val.endswith('汇总') or plat_val == '总计'):
            continue
        if r == 3:
            continue
        series_val = sheet.cell(row=r, column=2).value
        model_val = sheet.cell(row=r, column=3).value
        detail_rows.append((r, plat_val, series_val, model_val))

    for row_idx in range(start_row, max_row + 1):
        plat_cell = sheet.cell(row=row_idx, column=1)
        if plat_cell.value and isinstance(plat_cell.value, str):
            val = plat_cell.value
            if val.endswith('汇总'):
                sheet.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=3)
                b_val = sheet.cell(row=row_idx, column=2).value
                merged_cell = sheet.cell(row=row_idx, column=2)
                merged_cell.value = b_val
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row_idx, column=col)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            elif val == '总计':
                sheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
                merged_cell = sheet.cell(row=row_idx, column=1)
                merged_cell.value = '总计'
                merged_cell.alignment = Alignment(horizontal='center', vertical='center')
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row_idx, column=col)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    if detail_rows:
        platform_groups = defaultdict(list)
        for r, plat, series, model in detail_rows:
            platform_groups[plat].append((r, series, model))
        for plat, rows in platform_groups.items():
            start_r = rows[0][0]
            end_r = rows[-1][0]
            if len(rows) > 1:
                sheet.merge_cells(start_row=start_r, start_column=1,
                                  end_row=end_r, end_column=1)
                merged_cell = sheet.cell(row=start_r, column=1)
                merged_cell.value = plat

    if detail_rows:
        platform_groups = defaultdict(list)
        for r, plat, series, model in detail_rows:
            platform_groups[plat].append((r, series, model))
        for plat, rows in platform_groups.items():
            idx = 0
            while idx < len(rows):
                row_idx, series_val, model_val = rows[idx]
                if series_val is None:
                    idx += 1
                    continue
                end_idx = idx
                while end_idx < len(rows) and str(rows[end_idx][1]) == str(series_val):
                    end_idx += 1
                if end_idx - idx > 1:
                    start_r = rows[idx][0]
                    end_r = rows[end_idx - 1][0]
                    sheet.merge_cells(start_row=start_r, start_column=2,
                                      end_row=end_r, end_column=2)
                    merged_cell = sheet.cell(row=start_r, column=2)
                    merged_cell.value = series_val
                sub_rows = rows[idx:end_idx]
                j = 0
                while j < len(sub_rows):
                    r_j, _, model_val_j = sub_rows[j]
                    if model_val_j is None:
                        j += 1
                        continue
                    j2 = j
                    while j2 < len(sub_rows) and sub_rows[j2][2] == model_val_j:
                        j2 += 1
                    if j2 - j > 1:
                        start_r_model = sub_rows[j][0]
                        end_r_model = sub_rows[j2 - 1][0]
                        sheet.merge_cells(start_row=start_r_model, start_column=3,
                                          end_row=end_r_model, end_column=3)
                        merged_cell = sheet.cell(row=start_r_model, column=3)
                        merged_cell.value = model_val_j
                    j = j2
                idx = end_idx

    col_widths = {1: 22, 2: 18, 3: 12, 4: 14, 5: 16, 6: 16, 7: 16, 8: 14}
    for col, width in col_widths.items():
        if col <= max_col:
            sheet.column_dimensions[get_column_letter(col)].width = width
    for col in range(9, max_col + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 14

    debug_print("apply_merges 完成")


# ---------- HTML 生成 ----------
def df_to_html_with_merges(df: pd.DataFrame, title: str, font_size: int = FONT_SIZE_HTML) -> str:
    """将DataFrame转换为带合并单元格的HTML表格（含国内/国际标题）。

        Args:
            df: 分组统计DataFrame，应包含attrs（sorted_regions等）。
            title: 表格标题（如'国内'或'国际'）。
            font_size: 表格内字体大小。

        Returns:
            完整的HTML表格字符串。
        """
    debug_print(f"开始 df_to_html_with_merges for {title}")

    # ---- 新增预处理：将可能为NaN的字符串列转为空字符串，防止死循环 ----
    df = df.copy()
    for col in ['平台类别', '产品系列', '产品型号', '定制原因']:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str)
    # ----------------------------------------------------------------


    rows = df.to_dict('records')
    if not rows:
        return "<p>无数据</p>"

    cols = df.columns.tolist()
    sorted_regions = df.attrs.get('sorted_regions', [])
    region_data = df.attrs.get('region_data', {})

    std_col = next((c for c in cols if '标准化' in c and not any(c.startswith(r) for r in sorted_regions)), None)
    cfg_col = next((c for c in cols if '配置化' in c and not any(c.startswith(r) for r in sorted_regions)), None)
    cus_col = next((c for c in cols if '定制化' in c and not any(c.startswith(r) for r in sorted_regions)), None)
    area_col = '面积汇总'
    custom_col = '定制原因'
    total_col = '总计'

    thead_parts = ['<thead><tr>']
    for col_name in ['平台类别', '产品系列', '产品型号', '面积汇总', std_col, cfg_col, cus_col, custom_col]:
        thead_parts.append(f'<th rowspan="2">{col_name}</th>')
    for region in sorted_regions:
        info = region_data.get(region, {'排名': 0, '面积汇总': 0, '占比': 0})
        rank = info['排名']
        area = info['面积汇总']
        ratio = info['占比']
        display_text = f"TOP{rank:.0f} - {region}<br>{area:.2f}㎡ - ({ratio:.1f}%)"
        thead_parts.append(f'<th colspan="3">{display_text}</th>')
    thead_parts.append(f'<th rowspan="2">总计</th>')
    thead_parts.append('</tr><tr>')
    for region in sorted_regions:
        thead_parts.append('<th>标准化</th><th>配置化</th><th>定制化</th>')
    thead_parts.append('</tr></thead>')
    thead = ''.join(thead_parts)

    tbody_rows = []
    row_types = []
    for row in rows:
        plat = row['平台类别']
        if plat == '总计':
            row_types.append('total')
        elif isinstance(plat, str) and plat.endswith('汇总'):
            row_types.append('subtotal')
        elif plat == '':
            row_types.append('top_total')
        else:
            row_types.append('detail')

    i = 0
    total_rows = len(rows)

    while i < total_rows:
        row = rows[i]
        row_type = row_types[i]

        if row_type == 'top_total':
            tr = '<tr class="top-total">'
            tr += '<td colspan="3">总计</td>'
            area_val = row.get(area_col, 0)
            std_val = row.get(std_col, 0)
            cfg_val = row.get(cfg_col, 0)
            cus_val = row.get(cus_col, 0)
            custom_val = row.get(custom_col, '')
            tr += f'<td>{"" if area_val==0 else str(int(area_val))}</td>'
            tr += f'<td>{"" if std_val==0 else str(int(std_val))}</td>'
            tr += f'<td>{"" if cfg_val==0 else str(int(cfg_val))}</td>'
            tr += f'<td>{"" if cus_val==0 else str(int(cus_val))}</td>'
            tr += f'<td>{custom_val}</td>'
            for region in sorted_regions:
                std_col_r = f"{region}_标准化"
                cfg_col_r = f"{region}_配置化"
                cus_col_r = f"{region}_定制化"
                std_val_r = row.get(std_col_r, 0)
                cfg_val_r = row.get(cfg_col_r, 0)
                cus_val_r = row.get(cus_col_r, 0)
                tr += f'<td>{"" if std_val_r==0 else str(int(std_val_r))}</td>'
                tr += f'<td>{"" if cfg_val_r==0 else str(int(cfg_val_r))}</td>'
                tr += f'<td>{"" if cus_val_r==0 else str(int(cus_val_r))}</td>'
            total_val = row.get(total_col, 0)
            tr += f'<td>{"" if total_val==0 else str(int(total_val))}</td>'
            tr += '</tr>'
            tbody_rows.append(tr)
            i += 1
            continue

        if row_type in ('subtotal', 'total'):
            plat_val = row['平台类别']
            if row_type == 'subtotal':
                series_val = row['产品系列']
                css_class = 'subtotal'
            else:
                series_val = ''
                css_class = 'total'
            custom_val = ''
            area_val = row.get(area_col, 0)
            std_val = row.get(std_col, 0)
            cfg_val = row.get(cfg_col, 0)
            cus_val = row.get(cus_col, 0)
            total_val = row.get(total_col, 0)

            area_disp = '' if area_val == 0 else str(int(area_val))
            std_disp = '' if std_val == 0 else str(int(std_val))
            cfg_disp = '' if cfg_val == 0 else str(int(cfg_val))
            cus_disp = '' if cus_val == 0 else str(int(cus_val))
            total_disp = '' if total_val == 0 else str(int(total_val))

            tr = f'<tr class="{css_class}">'
            if row_type == 'total':
                tr += f'<td colspan="3">{plat_val}</td>'
            else:
                tr += f'<td>{plat_val}</td>'
                tr += f'<td colspan="2">{series_val}</td>'
            tr += f'<td>{area_disp}</td>'
            tr += f'<td>{std_disp}</td>'
            tr += f'<td>{cfg_disp}</td>'
            tr += f'<td>{cus_disp}</td>'
            tr += f'<td>{custom_val}</td>'
            for region in sorted_regions:
                std_col_r = f"{region}_标准化"
                cfg_col_r = f"{region}_配置化"
                cus_col_r = f"{region}_定制化"
                std_val_r = row.get(std_col_r, 0)
                cfg_val_r = row.get(cfg_col_r, 0)
                cus_val_r = row.get(cus_col_r, 0)
                tr += f'<td>{"" if std_val_r==0 else str(int(std_val_r))}</td>'
                tr += f'<td>{"" if cfg_val_r==0 else str(int(cfg_val_r))}</td>'
                tr += f'<td>{"" if cus_val_r==0 else str(int(cus_val_r))}</td>'
            tr += f'<td>{total_disp}</td>'
            tr += '</tr>'
            tbody_rows.append(tr)
            i += 1
            continue

        # 明细行
        curr_plat = row['平台类别']
        start_i = i
        while i < total_rows and row_types[i] == 'detail' and rows[i]['平台类别'] == curr_plat:
            i += 1
        end_i = i - 1
        plat_rowspan = end_i - start_i + 1

        j = start_i
        while j <= end_i:
            curr_series = rows[j]['产品系列']
            series_start = j
            while j <= end_i and rows[j]['产品系列'] == curr_series:
                j += 1
            series_end = j - 1
            series_rowspan = series_end - series_start + 1

            k = series_start
            while k <= series_end:
                curr_model = rows[k]['产品型号']
                model_start = k
                while k <= series_end and rows[k]['产品型号'] == curr_model:
                    k += 1
                model_end = k - 1
                model_rowspan = model_end - model_start + 1

                for idx in range(model_start, model_end + 1):
                    row_data = rows[idx]
                    if idx == model_start:
                        custom_val = row_data.get(custom_col, '')
                        plat_td = f'<td rowspan="{plat_rowspan}">{curr_plat}</td>' if idx == start_i else ''
                        series_td = f'<td rowspan="{series_rowspan}">{curr_series}</td>' if idx == series_start else ''
                        model_td = f'<td rowspan="{model_rowspan}">{curr_model}</td>' if idx == model_start else ''
                    else:
                        custom_val = ''
                        plat_td = ''
                        series_td = ''
                        model_td = ''

                    area_val = row_data.get(area_col, 0)
                    std_val = row_data.get(std_col, 0)
                    cfg_val = row_data.get(cfg_col, 0)
                    cus_val = row_data.get(cus_col, 0)
                    total_val = row_data.get(total_col, 0)

                    area_disp = '' if area_val == 0 else str(int(area_val))
                    std_disp = '' if std_val == 0 else str(int(std_val))
                    cfg_disp = '' if cfg_val == 0 else str(int(cfg_val))
                    cus_disp = '' if cus_val == 0 else str(int(cus_val))
                    total_disp = '' if total_val == 0 else str(int(total_val))

                    tr = '<tr>'
                    if plat_td:
                        tr += plat_td
                    if series_td:
                        tr += series_td
                    if model_td:
                        tr += model_td
                    tr += f'<td>{area_disp}</td>'
                    tr += f'<td>{std_disp}</td>'
                    tr += f'<td>{cfg_disp}</td>'
                    tr += f'<td>{cus_disp}</td>'
                    tr += f'<td>{custom_val}</td>'
                    for region in sorted_regions:
                        std_col_r = f"{region}_标准化"
                        cfg_col_r = f"{region}_配置化"
                        cus_col_r = f"{region}_定制化"
                        std_val_r = row_data.get(std_col_r, 0)
                        cfg_val_r = row_data.get(cfg_col_r, 0)
                        cus_val_r = row_data.get(cus_col_r, 0)
                        tr += f'<td>{"" if std_val_r==0 else str(int(std_val_r))}</td>'
                        tr += f'<td>{"" if cfg_val_r==0 else str(int(cfg_val_r))}</td>'
                        tr += f'<td>{"" if cus_val_r==0 else str(int(cus_val_r))}</td>'
                    tr += f'<td>{total_disp}</td>'
                    tr += '</tr>'
                    tbody_rows.append(tr)

                k = model_end + 1
            j = series_end + 1
        i = end_i + 1

    html = f"""
    <div class="table-container">
        <h2>{title}</h2>
        <table style="font-size: {font_size}px; border-collapse: collapse; width: 100%;">
            {thead}
            <tbody>
                {''.join(tbody_rows)}
            </tbody>
        </table>
    </div>
    """
    debug_print(f"df_to_html_with_merges for {title} 完成")
    return html


# ========== 工具函数 ==========
def get_download_folder() -> str:
    """获取当前用户的下载目录路径"""
    if os.name == 'nt':
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:
        return os.path.join(os.path.expanduser('~'), 'Downloads')


def rebuild_df_from_excel_bk(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """
    从已生成的分组统计 Excel 中读取数据，结合元数据 Sheet 重建 DataFrame（含 attrs）。
    注意：元数据 Sheet 被隐藏，但读取时不受影响。
    """
    debug_print(f"开始 rebuild_df_from_excel: {excel_path}, sheet={sheet_name}")
    meta_df = pd.read_excel(excel_path, sheet_name='元数据', header=None, index_col=0)
    if sheet_name == '国内':
        regions_key = 'sorted_regions_dom'
        data_key = 'region_data_dom'
    elif sheet_name == '国际':
        regions_key = 'sorted_regions_int'
        data_key = 'region_data_int'
    else:
        raise ValueError(f"不支持的 sheet_name: {sheet_name}，仅支持 '国内' 或 '国际'")

    try:
        sorted_regions = json.loads(meta_df.loc[regions_key, 1])
        region_data = json.loads(meta_df.loc[data_key, 1])
    except KeyError as e:
        raise KeyError(f"元数据中缺少键: {e}，请确认 Excel 是由本工具生成。")

    df_data = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=2)
    num_cols = 8 + 3 * len(sorted_regions) + 1
    df_data = df_data.iloc[:, :num_cols]

    std_col = '标准化'
    cfg_col = '配置化'
    cus_col = '定制化'
    base_cols = ['平台类别', '产品系列', '产品型号', '面积汇总', std_col, cfg_col, cus_col, '定制原因']
    region_cols = []
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            region_cols.append(f"{region}_{cat}")
    final_cols = base_cols + region_cols + ['总计']

    df_data.columns = final_cols

    numeric_cols = [c for c in final_cols if c not in ['平台类别', '产品系列', '产品型号', '定制原因']]
    for col in numeric_cols:
        df_data[col] = df_data[col].fillna(0)

    df_data['面积汇总'] = df_data[std_col] + df_data[cfg_col] + df_data[cus_col]

    df_data.attrs['sorted_regions'] = sorted_regions
    df_data.attrs['region_data'] = region_data
    df_data.attrs['region_area'] = {r: region_data[r]['面积汇总'] for r in sorted_regions}
    df_data.attrs['region_ratio'] = {r: region_data[r]['占比'] for r in sorted_regions}
    df_data.attrs['region_rank'] = {r: region_data[r]['排名'] for r in sorted_regions}

    debug_print(f"rebuild_df_from_excel 完成，行数：{len(df_data)}")
    return df_data


def rebuild_df_from_excel(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """
    从已生成的分组统计 Excel 中读取数据，结合元数据 Sheet 重建 DataFrame（含 attrs）。
    直接从 Excel 表头读取百分比列名，不依赖元数据计算。
    """
    debug_print(f"开始 rebuild_df_from_excel: {excel_path}, sheet={sheet_name}")

    # 1. 从元数据中读取区域信息（区域列表和区域统计数据）
    meta_df = pd.read_excel(excel_path, sheet_name='元数据', header=None, index_col=0)
    if sheet_name == '国内':
        regions_key = 'sorted_regions_dom'
        data_key = 'region_data_dom'
    elif sheet_name == '国际':
        regions_key = 'sorted_regions_int'
        data_key = 'region_data_int'
    else:
        raise ValueError(f"不支持的 sheet_name: {sheet_name}，仅支持 '国内' 或 '国际'")

    try:
        sorted_regions = json.loads(meta_df.loc[regions_key, 1])
        region_data = json.loads(meta_df.loc[data_key, 1])
    except KeyError as e:
        raise KeyError(f"元数据中缺少键: {e}，请确认 Excel 是由本工具生成。")

    # 2. 读取 Excel 数据（跳过前两行，即表头行）
    df_data = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=2)
    num_cols = 8 + 3 * len(sorted_regions) + 1
    df_data = df_data.iloc[:, :num_cols]

    # 3. 从 Excel 第一行读取前 8 列的列名（含百分比）
    header_row = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=1).iloc[0]
    base_cols_from_excel = [str(header_row[i]) for i in
                            range(8)]  # 前8列：平台类别、产品系列、产品型号、面积汇总、标准化(XX%)、配置化(XX%)、定制化(XX%)、定制原因

    # 4. 构建完整的列名列表
    # 前8列使用从 Excel 读取的名称
    region_cols = []
    for region in sorted_regions:
        for cat in ['标准化', '配置化', '定制化']:
            region_cols.append(f"{region}_{cat}")
    final_cols = base_cols_from_excel + region_cols + ['总计']

    df_data.columns = final_cols

    # 5. 向前填充合并单元格的空值（修复空值导致排版错乱）
    for col in ['平台类别', '产品系列', '产品型号']:
        if col in df_data.columns:
            df_data[col] = df_data[col].ffill()

    # 6. 数值列填充 0
    numeric_cols = [c for c in final_cols if
                    c not in ['平台类别', '产品系列', '产品型号', '定制原因']]
    for col in numeric_cols:
        df_data[col] = df_data[col].fillna(0)

    # 7. 确保“面积汇总”列正确（重新计算，保证一致性）
    # 查找标准化、配置化、定制化三列的实际名称（带百分比）
    std_col = base_cols_from_excel[4]
    cfg_col = base_cols_from_excel[5]
    cus_col = base_cols_from_excel[6]
    df_data['面积汇总'] = df_data[std_col] + df_data[cfg_col] + df_data[cus_col]

    # 8. 设置 attrs
    df_data.attrs['sorted_regions'] = sorted_regions
    df_data.attrs['region_data'] = region_data
    df_data.attrs['region_area'] = {r: region_data[r]['面积汇总'] for r in sorted_regions}
    df_data.attrs['region_ratio'] = {r: region_data[r]['占比'] for r in sorted_regions}
    df_data.attrs['region_rank'] = {r: region_data[r]['排名'] for r in sorted_regions}

    debug_print(f"rebuild_df_from_excel 完成，行数：{len(df_data)}")
    return df_data



# ========== GUI 应用程序类 ==========
class GUIApp:
    """分组统计工具图形界面主类"""

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        master.title("分组统计工具")
        master.geometry("800x600")
        master.resizable(True, True)

        self.data_file_path = tk.StringVar()
        self.excel_file_path = tk.StringVar()
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.btn_stat = None
        self.btn_convert = None
        self.btn_add_data = None
        self.btn_select_excel = None

        # 用于保存转换过程中读取的DataFrame（跨线程传递）
        self._converted_df_dom = None
        self._converted_df_int = None
        self._converted_excel_path = None

        self._create_widgets()

    def _create_widgets(self) -> None:
        # 顶部提示区域
        top_frame = tk.Frame(self.master)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        hint_label = tk.Label(top_frame, text="使用说明", font=('Arial', 12, 'bold'))
        hint_label.pack(anchor='w')

        hint_text = scrolledtext.ScrolledText(
            top_frame,
            wrap=tk.WORD,
            font=('Arial', 13),
            height=10,
            relief=tk.SUNKEN,
            borderwidth=2
        )
        hint_text.pack(fill=tk.BOTH, expand=True)

        info = (
            "【数据文件要求】\n"
            "请选择包含原始数据的 Excel 文件（.xlsx 或 .xls），该文件必须包含以下列：\n"
            "  - 区域         : 取值为 '国内' 或 '国际'，用于区分国内/国际统计。\n"
            "  - 销售大区     : 如华东、华南等，用于按大区统计。\n"
            "  - 省份/国家    : 省份或国家名称（仅用于展示）。\n"
            "  - 国家         : 国家名称（仅用于展示）。\n"
            "  - 总面积       : 数值型，面积数据。\n"
            "  - 产品系列     : 产品所属系列。\n"
            "  - 产品型号     : 具体型号。\n"
            "  - 产品主推类型（产品维度）: 取值为 'TOP'、'TOP+'、'NON-TOP' 或 '定制'。\n"
            "  - 平台类别     : 如 '平台A'、'平台B' 等。\n"
            "此外，建议包含 '订单数量' 列（非必需，用于统计）。\n"
            "Excel 兼容格式：.xlsx 或 .xls。\n\n"
            "【操作说明】\n"
            " 第一步：先生成Excel格式的统计数据 \n"
                "1.1. 点击 '添加数据文件' 按钮选择原始数据 Excel。\n"
                "1.2. 点击 '统计' 按钮生成分组统计 Excel 和 HTML 文件（如无定制原因修改需求，该文件可以直接使用），并弹出保存对话框。\n\n"
             " 第二步：选中修改后的Excel文件（格式不可做调整），将其转成html文件 \n"
                "2.1. 对于已修改好的的Excel（仅限本工具生成的 Excel 格式）文件，可使用下方的 '选择 Excel 文件' 按钮加载。\n"
                "2.2. 点击 'Excel 转 HTML' 按钮将选中的 Excel 文件转换为 HTML。\n\n"
            " ***重要提醒：统计生成的Excel文件内容可以修改，格式不能调整，否则无法转成Html文件。*** \n"
        )
        hint_text.insert(tk.END, info, )
        hint_text.config(state=tk.DISABLED)

        # 底部功能区
        bottom_frame = tk.Frame(self.master)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        # ----- 功能区1：统计 -----
        group1 = tk.LabelFrame(bottom_frame, text="数据统计", padx=5, pady=5)
        group1.pack(fill=tk.X, pady=5)

        row1 = tk.Frame(group1)
        row1.pack(fill=tk.X, pady=2)

        tk.Label(row1, text="数据文件:", width=10, anchor='w').pack(side=tk.LEFT)
        entry_data = tk.Entry(row1, textvariable=self.data_file_path, width=50, state='readonly')
        entry_data.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.btn_add_data = tk.Button(row1, text="添加数据文件", command=self._select_data_file)
        self.btn_add_data.pack(side=tk.LEFT, padx=2)

        self.btn_stat = tk.Button(row1, text="统计", command=self._run_statistics)
        self.btn_stat.pack(side=tk.LEFT, padx=2)

        # ----- 功能区2：Excel 转 HTML -----
        # ----- 功能区2：Excel 转 HTML -----
        group2 = tk.LabelFrame(bottom_frame, text="Excel 转 HTML", padx=5, pady=5)
        group2.pack(fill=tk.X, pady=5)

        row2 = tk.Frame(group2)
        row2.pack(fill=tk.X, pady=2)

        tk.Label(row2, text="Excel 文件:", width=10, anchor='w').pack(side=tk.LEFT)
        entry_excel = tk.Entry(row2, textvariable=self.excel_file_path, width=40, state='readonly')
        entry_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 将两个按钮放入独立的Frame，并靠右放置，避免被输入框挤压
        btn_frame = tk.Frame(row2)
        btn_frame.pack(side=tk.RIGHT, padx=5)

        self.btn_select_excel = tk.Button(btn_frame, text="选择 Excel 文件", width= 10,
                                          command=self._select_excel_file)
        self.btn_select_excel.pack(side=tk.LEFT, padx=2)

        self.btn_convert = tk.Button(btn_frame, text="Excel 转 HTML", width= 15,
                                     command=self._convert_excel_to_html)
        self.btn_convert.pack(side=tk.LEFT, padx=2, ipadx=10)

        # 状态栏
        status_bar = tk.Label(self.master, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _select_data_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择原始数据文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.data_file_path.set(file_path)

    def _select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择分组统计 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_file_path.set(file_path)

    def _run_statistics(self) -> None:
        data_path = self.data_file_path.get()
        if not data_path:
            messagebox.showerror("错误", "请先选择数据文件。")
            return

        self._set_buttons_state(False)
        self.status_var.set("正在处理统计任务...")
        debug_print("统计任务开始")

        def task():
            try:
                self._do_statistics(data_path)
            except Exception as e:
                self.master.after(0, lambda: self._show_error(f"统计失败：{e}"))
            finally:
                self.master.after(0, self._finish_task)

        threading.Thread(target=task, daemon=True).start()

    def _do_statistics(self, data_path: str) -> None:
        debug_print(f"读取原始数据: {data_path}")
        data = pd.read_excel(data_path)
        required_cols = ['区域', '销售大区', '省份/国家', '国家', '总面积',
                         '产品系列', '产品型号', '产品主推类型（产品维度）', '平台类别']
        missing = [c for c in required_cols if c not in data.columns]
        if missing:
            raise ValueError(f"数据文件缺少必需列：{', '.join(missing)}")

        data['主推分类'] = data['产品主推类型（产品维度）'].apply(classify_main_type)

        domestic = data[data['区域'] == '国内'].copy()
        international = data[data['区域'] == '国际'].copy()

        if domestic.empty and international.empty:
            raise ValueError("数据中既无 '国内' 也无 '国际' 记录。")

        df_dom = process_region_data(domestic) if not domestic.empty else None
        df_int = process_region_data(international) if not international.empty else None

        debug_print("统计完成，准备弹出保存对话框")
        self.master.after(0, lambda: self._save_dialog(df_dom, df_int))

    def _save_dialog(self, df_dom: Optional[pd.DataFrame], df_int: Optional[pd.DataFrame]) -> None:
        default_dir = get_download_folder()
        file_path = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            initialdir=default_dir,
            initialfile="分组统计结果.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not file_path:
            self.status_var.set("已取消保存")
            self._finish_task()
            return

        try:
            debug_print(f"保存 Excel: {file_path}")
            self._save_excel(file_path, df_dom, df_int)
            html_path = os.path.splitext(file_path)[0] + ".html"
            debug_print(f"生成 HTML: {html_path}")
            self._generate_html(html_path, df_dom, df_int)
            self.master.after(0, lambda: self._show_success(f"统计完成！\nExcel：{file_path}\nHTML：{html_path}"))
            self.master.after(0, lambda: webbrowser.open(html_path))
        except Exception as e:
            self.master.after(0, lambda: self._show_error(f"保存失败：{e}"))
        finally:
            self.master.after(0, self._finish_task)

    def _save_excel(self, file_path: str, df_dom: Optional[pd.DataFrame], df_int: Optional[pd.DataFrame]) -> None:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            if df_dom is not None:
                df_dom_excel = df_dom.copy()
                numeric_cols = [c for c in df_dom_excel.columns if c not in ['平台类别', '产品系列', '产品型号', '定制原因']]
                for col in numeric_cols:
                    df_dom_excel[col] = df_dom_excel[col].replace(0, None)
                df_dom_excel.to_excel(writer, sheet_name='国内', index=False, header=False, startrow=2)
            if df_int is not None:
                df_int_excel = df_int.copy()
                numeric_cols = [c for c in df_int_excel.columns if c not in ['平台类别', '产品系列', '产品型号', '定制原因']]
                for col in numeric_cols:
                    df_int_excel[col] = df_int_excel[col].replace(0, None)
                df_int_excel.to_excel(writer, sheet_name='国际', index=False, header=False, startrow=2)

            meta_data = {}
            if df_dom is not None:
                meta_data['sorted_regions_dom'] = df_dom.attrs['sorted_regions']
                meta_data['region_data_dom'] = df_dom.attrs['region_data']
            if df_int is not None:
                meta_data['sorted_regions_int'] = df_int.attrs['sorted_regions']
                meta_data['region_data_int'] = df_int.attrs['region_data']

            meta_df = pd.DataFrame({
                'key': list(meta_data.keys()),
                'value': [json.dumps(v) for v in meta_data.values()]
            })
            meta_df.to_excel(writer, sheet_name='元数据', index=False, header=False)

            workbook = writer.book
            if '元数据' in workbook.sheetnames:
                meta_sheet = workbook['元数据']
                meta_sheet.sheet_state = 'hidden'

            for sheet_name in ['国内', '国际']:
                if sheet_name not in workbook.sheetnames:
                    continue
                sheet = workbook[sheet_name]
                if sheet_name == '国内' and df_dom is not None:
                    sorted_regions = df_dom.attrs.get('sorted_regions', [])
                    region_data = df_dom.attrs.get('region_data', {})
                    col_names = df_dom.columns.tolist()
                elif sheet_name == '国际' and df_int is not None:
                    sorted_regions = df_int.attrs.get('sorted_regions', [])
                    region_data = df_int.attrs.get('region_data', {})
                    col_names = df_int.columns.tolist()
                else:
                    continue
                apply_merges(sheet, start_row=3, sorted_regions=sorted_regions,
                             region_data=region_data, col_names=col_names)

    def _generate_html(self, html_path: str, df_dom: Optional[pd.DataFrame], df_int: Optional[pd.DataFrame]) -> None:
        debug_print("生成HTML内容")
        dom_html = df_to_html_with_merges(df_dom, "国内", font_size=FONT_SIZE_HTML) if df_dom is not None else ""
        int_html = df_to_html_with_merges(df_int, "国际", font_size=FONT_SIZE_HTML) if df_int is not None else ""

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>分组统计结果</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h2 {{ text-align: center; color: #333; }}
                .table-container {{ margin-bottom: 40px; overflow-x: auto; }}
                table {{ border-collapse: collapse; width: 100%; margin: 0 auto; }}
                th, td {{ border: 1px solid #000; padding: 6px 12px; text-align: center; }}
                th {{ background-color: #4A90D9; color: white; font-weight: bold; }}
                .top-total td {{ background-color: #D9E1F2; font-weight: bold; }}
                .total td {{ background-color: #D9E1F2; font-weight: bold; }}
                .subtotal td {{ background-color: #E2EFDA; font-weight: bold; }}
            </style>
        </head>
        <body>
            {dom_html}
            {int_html}
        </body>
        </html>
        """
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        debug_print(f"HTML保存完成: {html_path}")

    def _convert_excel_to_html(self) -> None:
        excel_path = self.excel_file_path.get()
        if not excel_path:
            messagebox.showerror("错误", "请先选择 Excel 文件。")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("错误", "文件不存在。")
            return
        if not excel_path.endswith('.xlsx'):
            messagebox.showerror("错误", "仅支持 .xlsx 格式。")
            return

        self._set_buttons_state(False)
        self.status_var.set("正在读取数据...")
        debug_print("Excel转HTML任务开始")

        self._converted_df_dom = None
        self._converted_df_int = None
        self._converted_excel_path = excel_path

        def read_task():
            try:
                self._do_convert_read(excel_path)
            except Exception as e:
                self.master.after(0, lambda: self._show_error(f"读取数据失败：{e}"))
                self.master.after(0, self._finish_task)

        threading.Thread(target=read_task, daemon=True).start()

    def _do_convert_read(self, excel_path: str) -> None:
        debug_print(f"开始读取Excel: {excel_path}")
        df_dom = None
        df_int = None
        with pd.ExcelFile(excel_path) as xls:
            if '元数据' not in xls.sheet_names:
                raise ValueError("该 Excel 不是由本工具生成（缺少元数据 Sheet）。")
            if '国内' in xls.sheet_names:
                df_dom = rebuild_df_from_excel(excel_path, '国内')
            if '国际' in xls.sheet_names:
                df_int = rebuild_df_from_excel(excel_path, '国际')

        if df_dom is None and df_int is None:
            raise ValueError("Excel 中既无 '国内' 也无 '国际' 数据。")

        self._converted_df_dom = df_dom
        self._converted_df_int = df_int
        debug_print("数据读取完成，准备弹出保存对话框")
        self.master.after(0, self._show_convert_save_dialog)

    def _show_convert_save_dialog(self) -> None:
        """主线程：弹出保存对话框并启动后台生成HTML"""
        excel_path = self._converted_excel_path
        df_dom = self._converted_df_dom
        df_int = self._converted_df_int

        default_dir = get_download_folder()
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        html_path = filedialog.asksaveasfilename(
            title="保存 HTML 文件",
            initialdir=default_dir,
            initialfile=base_name + ".html",
            defaultextension=".html",
            filetypes=[("HTML 文件", "*.html")]
        )
        if not html_path:
            self.status_var.set("已取消保存")
            self._finish_task()
            return

        # 在后台线程中生成HTML并保存
        self.status_var.set("正在生成HTML...")
        debug_print(f"准备后台生成HTML: {html_path}")

        def save_task():
            try:
                self._generate_html(html_path, df_dom, df_int)
                self.master.after(0, lambda: self._show_success(f"HTML 已保存至：{html_path}"))
                self.master.after(0, lambda: webbrowser.open(html_path))
            except Exception as e:
                self.master.after(0, lambda: self._show_error(f"保存HTML失败：{e}"))
            finally:
                self.master.after(0, self._finish_task)

        threading.Thread(target=save_task, daemon=True).start()

    # ---------- 辅助方法 ----------
    def _set_buttons_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        if self.btn_add_data:
            self.btn_add_data.config(state=state)
        if self.btn_stat:
            self.btn_stat.config(state=state)
        if self.btn_select_excel:
            self.btn_select_excel.config(state=state)
        if self.btn_convert:
            self.btn_convert.config(state=state)

    def _finish_task(self) -> None:
        self._set_buttons_state(True)
        self.status_var.set("就绪")
        debug_print("任务结束")

    def _show_error(self, msg: str) -> None:
        messagebox.showerror("错误", msg)
        self.status_var.set("错误")

    def _show_success(self, msg: str) -> None:
        messagebox.showinfo("成功", msg)
        self.status_var.set("完成")


# ========== 主入口 ==========
def main() -> None:
    root = tk.Tk()
    app = GUIApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()