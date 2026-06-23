import openpyxl
import pandas as pd
import numpy as np

pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

# 定义映射函数
def classify_main_type(val):
    if val in ['TOP', 'TOP+']:    # TOP 或 TOP+ 视为标准化
        return '标准化'
    elif val == 'NON-TOP':
        return '配置化'
    elif val == '定制':
        return '定制化'
    else:
        return '其他'   # 兜底

def main():
    # 在下面的代码行中使用断点来调试脚本。
    data_path = 'data/接单明细V1.1.xlsx'
    data = pd.read_excel(data_path)
    print(data.columns)

    titles = ['创建日期', '区域', '销售大区',
       '省份/国家', '国家', '订单数量', '总面积', '订单推送金额-RMB', '签单金额（万元）',
       '产品系列', '产品型号', '产品间距', '产品主推类型（产品维度）', '修正后的产品线',
       '修正后的业务产品线', '品牌', '业务产品线', '旗舰类别', '平台类别', '定位']

    data = data[titles]

    data['主推分类'] = data['产品主推类型（产品维度）'].apply(classify_main_type)

    result_data = pd.DataFrame()

    grouped = data.groupby(['平台类别', '区域', '销售大区', '产品系列', '产品型号']).agg({
        '总面积': 'sum',
        '订单数量': 'sum'
    }).reset_index()

    # 合并分类信息（确保产品型号唯一）
    type_map = data[['产品型号', '产品主推类型（产品维度）', '主推分类']].drop_duplicates(subset='产品型号')
    grouped = grouped.merge(type_map, on='产品型号', how='left')

    # 现在 grouped 包含所有需要的列
    print(grouped.head())

    # 透视生成三列
    pivot_df = grouped.pivot_table(
        index=['平台类别', '产品系列', '产品型号'],
        columns='主推分类',
        values='总面积',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # 确保三列存在
    for cat in ['标准化', '配置化', '定制化']:
        if cat not in pivot_df.columns:
            pivot_df[cat] = 0

    pivot_df['面积汇总'] = pivot_df[['标准化', '配置化', '定制化']].sum(axis=1)

    final_df = pivot_df[
        ['平台类别', '产品系列', '产品型号', '面积汇总', '标准化', '配置化', '定制化']]
    final_df.columns = ['平台类别', '产品系列', '产品型号', '面积汇总', '标准化', '配置化',
                        '定制化']

    # --- 分类汇总示例 ---
    # 1. 全局按分类汇总面积
    # total_by_class = grouped.groupby('主推分类')['总面积'].sum().reset_index()
    # print("\n各分类总面积：\n", total_by_class)
    #
    # # 2. 按平台类别+分类汇总
    # by_platform = grouped.groupby(['平台类别', '主推分类'])['总面积'].sum().reset_index()
    # print("\n按平台类别和分类：\n", by_platform)
    #
    # # 3. 按区域+分类汇总
    # by_region = grouped.groupby(['区域', '主推分类'])['总面积'].sum().reset_index()
    # print("\n按区域和分类：\n", by_region)

    print(final_df)

    # 分 sheet 保存
    df_domestic = final_df[grouped['区域'] == '国内']
    df_international = final_df[grouped['区域'] == '国际']

    with pd.ExcelWriter('分组统计结果.xlsx', engine='openpyxl') as writer:
        df_domestic.to_excel(writer, sheet_name='国内', index=False)
        df_international.to_excel(writer, sheet_name='国际', index=False)

    print("保存完成！")

    # 保存成一个excel文件，中两个sheet页，按照区域国内一个sheet，国际一个sheet



# 按装订区域中的绿色按钮以运行脚本。
if __name__ == '__main__':
    main()
