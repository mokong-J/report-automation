import re
import sys
import textwrap
from datetime import datetime
from pathlib import Path

import pandas as pd
from charset_normalizer import from_path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.text import CharacterProperties

# 质量报告行由各处理函数分别追加，main 里统一写盘（跨函数共享，故为模块级）
report_list = ["========== 数据质量报告 =========="]


def norm(s):
    """
    标准化字段名称：去首尾空格、移除空格/下划线/短横线、删除括号以及括号内内容、转小写
    例：金额（元—）-> 金额；转小写让列名匹配不区分大小写
    :param s: 输入字符串
    :return: 标准化后的字符串
    """
    s = s.strip()
    s = re.sub(r"[\s_-]", "", s)
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = s.lower()
    return s


def load_and_merge(data_dir, std_cols, alias):
    """
    遍历目录读取全部销售文件，统一列名后纵向合并。
    单个文件读取失败只登记不中断，保证坏文件不拖垮整批数据。
    :param data_dir: 脏数据目录 Path
    :param std_cols: 标准列名列表，决定输出列顺序
    :param alias: 标准名 -> 别名列表（元素已标准化）
    :return: 合并后的 DataFrame（可能含缺失，由 clean_data 统一处理）
    """
    failed_files = []
    new_data = pd.DataFrame()
    file_count = 0
    # 遍历所有Excel文件和csv文件
    for file_path in [*data_dir.glob("*.xls*"), *data_dir.glob("*.csv")]:
        if file_path.name.startswith("~$"):
            continue
        try:
            if file_path.suffix == ".csv":
                # csv 文件编码未知，先探测再按探测结果读取
                encoding = from_path(file_path).best().encoding
                df = pd.read_csv(file_path, encoding=encoding)
            else:
                df = pd.read_excel(file_path, engine="calamine")
        except Exception as e:
            print(f"错误：{file_path}读取失败，错误信息：{e}")
            failed_files.append(f"文件：{file_path.name} - 原因：{e}")
            continue

        if df.empty:
            print(f"错误：{file_path}为空，跳过")
            failed_files.append(f"文件：{file_path.name} - 原因：文件为空")
            continue

        # 统一列名：先标准化列名，再按别名表重命名
        # 每个标准名只取第一个命中的列（按字段名匹配、不按列位置匹配）
        df.columns = [norm(col) for col in df.columns]
        rename_map = {}
        for key, values in alias.items():
            for col in df.columns:
                if col in values:
                    rename_map[col] = key
                    break
        df = df.rename(columns=rename_map)

        # 字段缺失明说但不中断：缺的标准列由 reindex 补 NaN，后续统一按缺失行剔除
        missing_cols = [col for col in std_cols if col not in df.columns]
        if missing_cols:
            print(f"警告：{file_path.name} 缺少字段{','.join(missing_cols)}")

        df = df.reindex(columns=std_cols)

        new_data = pd.concat([new_data, df], ignore_index=True)

        file_count += 1
    report_list.append(f"成功获取{file_count}份数据，共{new_data.shape[0]}条数据")
    failed_order = (
        textwrap.indent("\n".join(failed_files), "  ") if failed_files else "无"
    )
    report_list.append(f"获取失败清单：\n{failed_order}")
    return new_data


def clean_data(df, std_cols):
    """
    清洗合并后的数据：日期解析、数值转换、剔缺失、地区标准化、去重。
    每一步的剔除量都计入质量报告，最终行数做自洽校验，不平时终止。
    :param df: load_and_merge 的合并结果
    :param std_cols: 标准列名，缺失剔除与去重都以这些字段为准
    :return: 清洗后的 DataFrame
    """
    total_rows = df.shape[0]
    # 日期格式统一化：先把「年 / 月 / . / /」替换为「-」并去掉「日」，
    # 再按固定格式解析，解析失败的置为 NaT（不中断，数量计入质量报告）
    # 注意「日」必须删除而非替换为「-」，否则尾缀多出分隔符导致解析失败
    df["Date"] = pd.to_datetime(
        df["Date"]
        .astype(str)
        .str.strip()
        .str.replace(r"[年月./]", "-", regex=True)
        .str.replace(r"[日]", "", regex=True),
        errors="coerce",
        format="%Y-%m-%d",
    )
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")
    date_na_rows = df["Date"].isna().sum()
    na_date_str = f"日期格式统一化：共{date_na_rows}个日期格式错误"
    print(na_date_str)
    report_list.append(f"  日期无法解析：{date_na_rows}行")

    # 数值列清洗：混入的「暂无 / - / N/A」等文本统一转为 NaN
    for col in ["Amount", "Quantity"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    na_amount_rows = df["Amount"].isna().sum()
    na_quantity_rows = df["Quantity"].isna().sum()
    report_list.append(f"  金额无法解析：{na_amount_rows}行（与其他原因有重叠）")
    report_list.append(f"  数量无法解析：{na_quantity_rows}行（与其他原因有重叠）")

    # 只剔除关键字段缺失的行（任一标准列缺失即剔除），不做全表 dropna
    rows_before_dropna = df.shape[0]
    df = df.dropna(subset=std_cols).copy()
    rows_after_dropna = df.shape[0]
    rows_na = rows_before_dropna - rows_after_dropna
    print(f"删除缺失值：共 {rows_na} 行缺失值（任一关键字段缺失的行）")
    report_list.append(f"  缺失值剔除：{rows_na}行（任一关键字段缺失的合计）")

    # 数量列用可空整数类型，避免输出成 17.0；缺失行已剔除，此处转换不会失败
    df["Quantity"] = df["Quantity"].astype("Int64")

    # 地区/产品标准化：去空格与括号注记，让「华东（大区）」并入「华东」
    # 放在去重之前，归一化后相同的行也能被一并判重
    for col in ["Region", "Product"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .str.replace(r"[\s_-]", "", regex=True)
            .str.replace(r"[（(].*?[)）]", "", regex=True)
        )

    # 去重：所有标准字段都相同才算重复，保留第一条
    rows_dup = df.duplicated(subset=std_cols).sum()
    df = df.drop_duplicates(subset=std_cols, keep="first")

    print(f"去重：共 {rows_dup} 行重复数据")
    report_list.append(f"  重复行剔除：{rows_dup}行")

    rows_cleaning = rows_dup + rows_na

    if total_rows != df.shape[0] + rows_cleaning:
        sys.exit(
            f"合并后数据行数与原始数据行数不一致，终止处理，原始数据行数：{total_rows}，"
            f"合并后数据行数：{df.shape[0]}，缺失值行数：{rows_na}，去重行数：{rows_dup}"
        )
    report_list.append(
        f"输出行数：{df.shape[0]}",
    )
    report_list.append(
        f"自洽校验：{rows_dup} + {rows_na} + {df.shape[0]} = {total_rows} ✓",
    )

    return df


def display_width(s):
    """按 Excel 显示宽度计算字符串宽度：中文等全角字符算 2，其余算 1"""
    width_s = 0
    for c in str(s):
        if ord(c) > 127:
            width_s += 2
        else:
            width_s += 1
    return width_s


def write_df(df, ws):
    """
    把 DataFrame 写入工作表：表头带样式，数据居中，列宽按内容自适应，冻结首行。
    :param df: 要写入的数据
    :param ws: 目标工作表
    """
    # 样式对象只建一次循环复用：逐格新建在大数据量时会明显拖慢写入
    cell_font = Font(bold=True, color="FFFFFF")
    cell_fill = PatternFill(fill_type="solid", start_color="4472C4")
    cell_alignment = Alignment(horizontal="center")
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = cell_font
        cell.fill = cell_fill
        cell.alignment = cell_alignment

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, col_value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=col_value)
            cell.alignment = cell_alignment

    for col_idx in range(1, ws.max_column + 1):
        widest = 0
        for row_idx in range(1, ws.max_row + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is not None:
                widest = max(widest, display_width(v))
        # 宽度 = 最宽内容 + 2 余量，但封顶 50，防超长脏值撑爆版面
        ws.column_dimensions[get_column_letter(col_idx)].width = min(widest + 2, 50)
    ws.freeze_panes = "A2"


def build_summary(df):
    """
    从清洗后的数据计算简报指标并拼装结论文字。
    只负责算与拼，写入由 write_summary 完成；返回总额供对账用。
    :param df: 清洗后且已含 Month 列的数据（金额单位：万元）
    :return: (结论文字列表, 总销售额)
    """
    total = df["Amount"].sum()
    quantity = df["Quantity"].sum()

    # 各地区销售额降序，取前3；最佳月份取合计最高的月份
    by_region = df.groupby("Region")["Amount"].sum().sort_values(ascending=False)
    top3 = by_region.head(3)

    by_month = df.groupby("Month")["Amount"].sum()
    best_month = by_month.idxmax()
    best_month_amount = by_month.max()
    lines = [
        f'报告期：{df["Date"].min()} ~ {df["Date"].max()}',
        f"销售总额：{total:.2f} 万元，总销量 {quantity:,.0f} 件",
    ]

    for rank, (region, amount) in enumerate(top3.items(), start=1):
        lines.append(
            f"Top{rank} 地区：{region}，销售额 {amount:.2f} 万元，占总额 {amount / total:.1%}"
        )

    lines.append(
        f"最佳月份：{best_month}，销售额 {best_month_amount:.2f} 万元，"
        f"占总额 {best_month_amount / total:.1%}",
    )

    return lines, total


def write_summary(ws, lines, total, pivot):
    """
    写入「简报」工作表，写入前先与透视表自动对账。
    对账不通过立即断言失败终止，保证简报数字与透视表永远一致。
    :param ws: 预先建好的「简报」工作表
    :param lines: build_summary 生成的结论文字
    :param total: 简报总销售额（万元）
    :param pivot: reset_index 后的透视表（末行末列即总计）
    """
    # 自动对账：iloc[-1, -1] 即 margins 生成的 Total 行 × Total 列
    pivot_total = pivot.iloc[-1, -1]
    assert (
        abs(total - pivot_total) < 0.01
    ), f"数字对账失败：简报总额 {total} ≠ 透视表合计 {pivot_total}"

    ws["B2"] = "销售周报简报"
    ws["B2"].font = Font(bold=True, size=14)

    for row_idx, text in enumerate(lines, start=4):
        cell = ws.cell(row=row_idx, column=2, value=text)
        cell.font = Font(size=11)

        # 合并横向单元格防止文字溢出，合并后只有左上角能赋值
        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=8)

    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["A"].width = 2


def _bold_title(chart):
    """图表标题加粗放大，两张图共用：sz 单位是 1/100 磅，1400 即 14 磅"""
    chart.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1400, b=True)


def _light_gridlines():
    """浅灰虚线横向网格线，两张图共用同一配置；w 单位 EMU，9525 约等于 0.75 磅"""
    return ChartLines(
        spPr=GraphicalProperties(
            ln=LineProperties(solidFill="D9D9D9", w=9525, prstDash="dash")
        )
    )


def add_charts(ws_data, ws_chart):
    """
    在图表工作表放置两张原生图表，数据引用透视表区域（不复制数据）。
    柱状图取透视表 Total 行（各地区合计），折线图取各月份行（各地区走势）。
    :param ws_data: 透视表工作表，作为图表数据源
    :param ws_chart: 图表工作表，作为图表摆放位置
    """
    last_row = ws_data.max_row
    total_col = ws_data.max_column

    # 柱状图：各地区销售总额，横轴=地区，纵轴=各地区销售总额
    bar = BarChart()
    bar.type = "col"
    bar.title = "各地区销售总额"
    _bold_title(bar)
    bar.x_axis.title = "地区"
    bar.y_axis.title = "销售额(万元)"
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.y_axis.number_format = "#,##0"
    bar.legend.position = "b"
    # 原 15x9 画布过小，自动布局空间不足时轴标题被挤进绘图区、与刻度/图例重叠
    bar.width = 20
    bar.height = 12
    # 网格线改浅灰细线；不要网格线时改为赋 None
    bar.y_axis.majorGridlines = _light_gridlines()

    # 手动绘制绘图区，防止标题等重合
    bar.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner",  # inner=只框绘图区（不含刻度标签），推荐
            xMode="edge",  # x/y 按"到图表区左/上边缘的距离"计算
            yMode="edge",
            x=0.08,  # 左上角：左边留 8% 给Y轴，上边留 11% 给图表标题
            y=0.11,
            w=0.85,  # 宽占 85%、高占 70%，底部 0.81 留给刻度/标题/图例
            h=0.7,
        )
    )

    # 数据：取 Total 行的地区列（不含 Total 列）
    data = Reference(
        ws_data, min_col=2, min_row=last_row, max_col=total_col - 1, max_row=last_row
    )
    bar.add_data(data, from_rows=True)

    # 分类轴取表头行的地区名
    cats = Reference(ws_data, min_col=2, min_row=1, max_col=total_col - 1, max_row=1)
    bar.set_categories(cats)
    ws_chart.add_chart(bar, "B2")

    # 折线图：月度销售额走势，横轴=月份，纵轴=销售额，每个地区一条线
    line = LineChart()
    line.title = "月度销售额走势"
    _bold_title(line)
    # majorTickMark=in 刻度线朝内；tickLblSkip=1 逐个显示月份标签不抽稀；
    # crossBetween=midCat 让数据点对齐刻度线中间而非骑在格线上
    line.x_axis.title = "月份"
    line.y_axis.title = "销售额(万元)"
    line.x_axis.delete = False
    line.y_axis.delete = False
    line.x_axis.majorTickMark = "in"
    line.y_axis.number_format = "#,##0"
    line.x_axis.tickLblSkip = 1
    line.x_axis.lblAlgn = "ctr"
    line.y_axis.crossBetween = "midCat"
    # 图例移到右侧：底部斜排的月份标签与图例争抢同一片空间导致重叠
    line.legend.position = "r"
    # 原 15x9 画布过小，加大后 X 轴标题、斜排月份标签、图例各有足够空间
    line.width = 20
    line.height = 12
    # 网格线改浅灰细线；不要网格线时改为赋 None
    line.y_axis.majorGridlines = _light_gridlines()

    # 手动绘制绘图区域，防止标题等重合
    line.layout = Layout(
        manualLayout=ManualLayout(
            layoutTarget="inner",
            xMode="edge",
            yMode="edge",
            x=0.08,
            y=0.12,
            w=0.75,
            h=0.75,
        )
    )

    # 数据：含表头行，titles_from_data 让地区名自动成为图例
    data = Reference(
        ws_data, min_col=2, min_row=1, max_col=total_col - 1, max_row=last_row - 1
    )
    line.add_data(data, titles_from_data=True)

    # 分类轴取 A 列的月份，不含表头和 Total 行
    cats = Reference(ws_data, min_col=1, min_row=2, max_col=1, max_row=last_row - 1)
    line.set_categories(cats)

    ws_chart.add_chart(line, "N2")
    # 左侧窄边距，图表不贴着工作表边缘
    ws_chart.column_dimensions["A"].width = 2


def main():
    """主流程：读脏数据 → 清洗 → 写四个 sheet（简报/原数据/透视表/图表）→ 保存 → 落盘质量报告"""
    date_str = datetime.now().strftime("%Y%m%d")
    # 输出文件名带运行当天日期，保留多天报表互不覆盖
    output_path = Path(__file__).parent / "output" / f"销售周报_{date_str}.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    report_path = Path(__file__).parent / "docs" / "report.txt"
    report_path.parent.mkdir(parents=True, exist_ok=True)

    # 标准列名，用于合并时的重命名与输出列顺序
    std_cols = ["Date", "Region", "Product", "Amount", "Quantity"]

    # 字段别名表，处理不同文件字段名不一致的情况（标准名 -> 别名列表）
    alias = {
        "Date": [
            "日期",
            "销售日期",
            "日期时间",
            "订单日期",
            "统计日期",
            "date",
            "order_date",
            "sale_date",
        ],
        "Region": [
            "地区",
            "销售区域",
            "区域",
            "大区",
            "省份",
            "城市",
            "region",
            "area",
            "city",
            "province",
        ],
        "Product": [
            "产品",
            "商品名称",
            "产品名称",
            "商品",
            "品类",
            "product",
            "item",
            "goods",
            "sku",
        ],
        "Amount": [
            "销售额",
            "销售金额",
            "金额",
            "金额(元)",
            "营业额",
            "amount",
            "sales",
            "sale_amount",
            "revenue",
            "gmv",
        ],
        "Quantity": [
            "数量",
            "销量",
            "销售数量",
            "件数",
            "qty",
            "quantity",
            "count",
            "units",
            "volume",
        ],
    }

    # 字段别名统一化，用于匹配标准化后的文件字段名
    alias = {key: [norm(value) for value in values] for key, values in alias.items()}
    file_dir = Path(__file__).parent / "dirty_data"
    if not file_dir.exists():
        sys.exit(f"错误：{file_dir}目录不存在")

    df = load_and_merge(file_dir, std_cols, alias)

    if df.empty:
        sys.exit("没有成功获取文件，终止处理")

    df = clean_data(df, std_cols)
    # 金额全局换算为万元并保留两位小数：原数据/透视表/图表/简报共用同一口径，
    # 若只在简报换算，会造成各 sheet 数字看起来对不上
    df["Amount"] = (df["Amount"] / 10000).round(2)

    # 最终 sheet 顺序：简报、原数据、透视表、图表（简报用 index=0 插到最前，打开先见结论）
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "原数据"
    ws2 = wb.create_sheet("透视表")
    ws3 = wb.create_sheet("图表")
    ws4 = wb.create_sheet("简报", 0)
    # 简报和图表不是数据表格，隐藏网格线版面更干净
    ws3.sheet_view.showGridLines = False
    ws4.sheet_view.showGridLines = False

    # 写入原数据
    write_df(df, ws1)

    # 制作透视表：Month 取日期前 7 位（YYYY-MM）作为行维度
    df["Month"] = df["Date"].astype(str).str[:7]
    pivot = pd.pivot_table(
        df,
        index="Month",
        columns="Region",
        values="Amount",
        aggfunc="sum",
        fill_value=0,  # 空值补 0
        margins=True,  # 自动加合计行列
        margins_name="Total",
    )

    pivot = pivot.reset_index()
    # 写入透视表
    write_df(pivot, ws2)

    # 添加图表
    add_charts(ws2, ws3)

    # 制作简报
    lines, total = build_summary(df)
    write_summary(ws4, lines, total, pivot)

    try:
        wb.save(output_path)
    except PermissionError:
        sys.exit("写入失败：输出文件被占用，请关闭 Excel 后重试")

    report = "\n".join(report_list)
    # print(report)
    with open(report_path, "w", encoding="utf-8") as f:
        f.writelines(report + "\n\n")


if __name__ == "__main__":
    main()
