import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "maison TOI ガントチャート"

# ===== スタイル定義 =====
header_font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
header_fill = PatternFill(start_color="2D2D2D", end_color="2D2D2D", fill_type="solid")
category_font = Font(name="Arial", bold=True, size=10)
category_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
task_font = Font(name="Arial", size=9)
thin_border = Border(
    left=Side(style="thin", color="DDDDDD"),
    right=Side(style="thin", color="DDDDDD"),
    top=Side(style="thin", color="DDDDDD"),
    bottom=Side(style="thin", color="DDDDDD"),
)

# カラーパレット（カテゴリごと）
colors = {
    "SNS・サイト": PatternFill(start_color="A8D8EA", end_color="A8D8EA", fill_type="solid"),  # 水色
    "ロンドン": PatternFill(start_color="D4A5E5", end_color="D4A5E5", fill_type="solid"),  # 紫
    "EC": PatternFill(start_color="FFD3B6", end_color="FFD3B6", fill_type="solid"),  # オレンジ
    "PR": PatternFill(start_color="FDFFAB", end_color="FDFFAB", fill_type="solid"),  # 黄
    "海外展開": PatternFill(start_color="B5EAD7", end_color="B5EAD7", fill_type="solid"),  # 緑
    "伊勢丹": PatternFill(start_color="FF9AA2", end_color="FF9AA2", fill_type="solid"),  # ピンク
}
milestone_fill = PatternFill(start_color="2D2D2D", end_color="2D2D2D", fill_type="solid")

# ===== 週の定義（4/7〜10/19） =====
start_date = datetime(2026, 4, 6)  # 週の始まり（月曜）
weeks = []
current = start_date
end_date = datetime(2026, 10, 19)
while current <= end_date:
    weeks.append(current)
    current += timedelta(weeks=1)

# ===== タスクデータ =====
# (カテゴリ, タスク名, 開始週index, 終了週index, カテゴリキー, マイルストーン)
tasks = [
    # ヘッダー行: カテゴリ
    ("cat", "SNS・サイト立ち上げ（4月）", None, None, None, None),
    ("task", "Instagramアカウント開設・プロフィール整備", 0, 0, "SNS・サイト", False),
    ("task", "ブランドサイト設計・コンテンツ準備", 0, 1, "SNS・サイト", False),
    ("task", "ブランドサイト構築（Claude Code）", 0, 1, "SNS・サイト", False),
    ("task", "Instagram初期投稿準備（9枚グリッド等）", 0, 1, "SNS・サイト", False),
    ("task", "★ サイト＆Instagram公開", 1, 1, "SNS・サイト", True),
    ("task", "Instagram運用開始（定期投稿）", 2, 27, "SNS・サイト", False),

    ("cat", "ロンドン活動（5〜7月）", None, None, None, None),
    ("task", "ロンドンのカフェ・セレクトショップ・ギャラリー調査", 2, 5, "ロンドン", False),
    ("task", "DMテンプレート・ルックブック準備", 3, 4, "ロンドン", False),
    ("task", "DM送付・アポ取り", 5, 12, "ロンドン", False),
    ("task", "展示イベント開催・交渉", 8, 14, "ロンドン", False),

    ("cat", "EC（6月〜）", None, None, None, None),
    ("task", "EC構築（日本向け）", 5, 8, "EC", False),
    ("task", "★ EC日本オープン（仮）", 8, 8, "EC", True),
    ("task", "EC運用・改善", 9, 27, "EC", False),

    ("cat", "PR・プレス（8〜9月）", None, None, None, None),
    ("task", "PR戦略策定・メディアリスト作成", 17, 19, "PR", False),
    ("task", "プレスリリース作成", 19, 20, "PR", False),
    ("task", "メディア・インフルエンサーへのアプローチ", 20, 23, "PR", False),
    ("task", "プレスサンプル貸出対応", 21, 25, "PR", False),

    ("cat", "海外展開（8〜9月）", None, None, None, None),
    ("task", "コペンハーゲン ギャラリー調査・DM送付", 17, 21, "海外展開", False),
    ("task", "パリ ギャラリー調査・DM送付", 17, 21, "海外展開", False),
    ("task", "海外展示交渉・スケジュール調整", 20, 25, "海外展開", False),

    ("cat", "伊勢丹ポップアップ（10月）", None, None, None, None),
    ("task", "伊勢丹 契約・条件詰め", 9, 14, "伊勢丹", False),
    ("task", "ポップアップ什器・ディスプレイ企画", 17, 22, "伊勢丹", False),
    ("task", "在庫・商品準備", 20, 25, "伊勢丹", False),
    ("task", "販促物制作（DM・フライヤー等）", 22, 25, "伊勢丹", False),
    ("task", "SNS告知・集客施策", 23, 26, "伊勢丹", False),
    ("task", "★ 伊勢丹ポップアップ開催", 27, 27, "伊勢丹", True),
]

# ===== ヘッダー行 =====
# 行1: 月表示
# 行2: 週表示
# 列A: カテゴリ/タスク名, 列B: ステータス, 列C以降: 週

ws.merge_cells("A1:B1")
ws["A1"].value = "maison TOI"
ws["A1"].font = Font(name="Arial", bold=True, size=14)
ws["A1"].alignment = Alignment(vertical="center")
ws.row_dimensions[1].height = 30

# 行2: 空行
ws.row_dimensions[2].height = 5

# 行3: 月ヘッダー
row_month = 3
row_week = 4

ws["A3"].value = "タスク"
ws["A3"].font = header_font
ws["A3"].fill = header_fill
ws["A3"].alignment = Alignment(vertical="center")
ws["B3"].value = "ステータス"
ws["B3"].font = header_font
ws["B3"].fill = header_fill
ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

ws["A4"].fill = header_fill
ws["B4"].fill = header_fill

# 月ごとにマージ
months = {}
for i, week in enumerate(weeks):
    col = i + 3  # C列から
    month_key = week.strftime("%Y年%-m月")
    if month_key not in months:
        months[month_key] = []
    months[month_key].append(col)

for month_name, cols in months.items():
    start_col = cols[0]
    end_col = cols[-1]
    if start_col == end_col:
        ws.cell(row=row_month, column=start_col).value = month_name
    else:
        ws.merge_cells(start_row=row_month, start_column=start_col,
                        end_row=row_month, end_column=end_col)
        ws.cell(row=row_month, column=start_col).value = month_name
    ws.cell(row=row_month, column=start_col).font = header_font
    ws.cell(row=row_month, column=start_col).fill = header_fill
    ws.cell(row=row_month, column=start_col).alignment = Alignment(horizontal="center")

    for c in cols:
        ws.cell(row=row_month, column=c).fill = header_fill

# 週番号
for i, week in enumerate(weeks):
    col = i + 3
    ws.cell(row=row_week, column=col).value = week.strftime("%-m/%-d")
    ws.cell(row=row_week, column=col).font = Font(name="Arial", size=8, color="FFFFFF")
    ws.cell(row=row_week, column=col).fill = header_fill
    ws.cell(row=row_week, column=col).alignment = Alignment(horizontal="center")

# ===== タスク行 =====
row = 5
for item in tasks:
    if item[0] == "cat":
        ws.cell(row=row, column=1).value = item[1]
        ws.cell(row=row, column=1).font = category_font
        ws.cell(row=row, column=1).fill = category_fill
        ws.cell(row=row, column=2).fill = category_fill
        for c in range(3, 3 + len(weeks)):
            ws.cell(row=row, column=c).fill = category_fill
        ws.row_dimensions[row].height = 22
    else:
        _, name, start_w, end_w, cat_key, is_milestone = item
        ws.cell(row=row, column=1).value = f"  {name}"
        ws.cell(row=row, column=1).font = task_font
        ws.cell(row=row, column=1).alignment = Alignment(vertical="center")
        ws.cell(row=row, column=2).value = "未着手"
        ws.cell(row=row, column=2).font = Font(name="Arial", size=8, color="999999")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 24

        if start_w is not None and end_w is not None:
            fill = milestone_fill if is_milestone else colors.get(cat_key, PatternFill())
            for w in range(start_w, end_w + 1):
                col = w + 3
                if col < 3 + len(weeks):
                    ws.cell(row=row, column=col).fill = fill

    # 罫線
    for c in range(1, 3 + len(weeks)):
        ws.cell(row=row, column=c).border = thin_border

    row += 1

# ===== 凡例 =====
row += 1
ws.cell(row=row, column=1).value = "凡例："
ws.cell(row=row, column=1).font = Font(name="Arial", bold=True, size=9)
row += 1
for cat_name, fill in colors.items():
    ws.cell(row=row, column=1).value = f"  {cat_name}"
    ws.cell(row=row, column=1).font = Font(name="Arial", size=9)
    ws.cell(row=row, column=2).fill = fill
    row += 1
ws.cell(row=row, column=1).value = "  ★ マイルストーン"
ws.cell(row=row, column=1).font = Font(name="Arial", size=9)
ws.cell(row=row, column=2).fill = milestone_fill
row += 1

# ===== 列幅 =====
ws.column_dimensions["A"].width = 45
ws.column_dimensions["B"].width = 10
for i in range(len(weeks)):
    ws.column_dimensions[get_column_letter(i + 3)].width = 5.5

# ===== フリーズ =====
ws.freeze_panes = "C5"

# 保存
filepath = "/Users/yayoi/clawd/maison_TOI_gantt.xlsx"
wb.save(filepath)
print(f"保存完了: {filepath}")
