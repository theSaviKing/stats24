# %%
import pandas as pd
import xlsxwriter as xlsx
import json as j

pd.set_option("future.no_silent_downcasting", True)
xl = pd.ExcelFile("Season 9.xlsx")

# %%
raw_sheet_names = [str(x) for x in xl.sheet_names[3:]]
raw_sheets: dict[str, pd.DataFrame] = dict()
for sheet_name in raw_sheet_names:
    sheet = (
        pd.read_excel(xl, sheet_name=sheet_name, header=None)
        .drop([0, 1, 2, 3])
        .drop(columns=[4, 5, 7, 8])
        .fillna({1: 0, 2: 0, 3: 0, 6: 0})
        .rename(columns={0: "Player", 1: "GT", 2: "GC", 3: "DP", 6: "Sub"})
    )
    sheet = sheet[
        ~sheet.Player.str.contains("Monday") & ~sheet.Player.str.contains("Tuesday")
    ]
    raw_sheets[sheet_name] = sheet


# %%
# def extract(
#     data: dict[int | str, pd.DataFrame], replacements: dict[str, str]
# ) -> dict[str, dict[str, dict[str, dict[str, int]]]]:
#     compiled = dict()
#     for sheet_name, sheet in data.items():
#         compiled[sheet_name] = dict()
#         current_team = str()
#         for row in sheet.itertuples():
#             pl = str(row.Player).strip()
#             if pl in replacements.keys():
#                 pl = replacements[pl]
#             if pl.isupper():
#                 current_team = pl
#                 continue
#             if current_team not in compiled[sheet_name].keys():
#                 compiled[sheet_name][current_team] = dict()
#             if pl in compiled[sheet_name][current_team].keys():
#                 compiled[sheet_name][current_team][pl]["GC"] += row.GC
#                 compiled[sheet_name][current_team][pl]["GT"] += row.GT
#                 compiled[sheet_name][current_team][pl]["DP"] += row.DP
#             else:
#                 compiled[sheet_name][current_team][pl] = {
#                     "GC": row.GC,
#                     "GT": row.GT,
#                     "DP": row.DP,
#                     "Sub": bool(row.Sub),
#                 }
#         for team in compiled[sheet_name].keys():
#             compiled[sheet_name][team] = {
#                 k: v
#                 for k, v in sorted(
#                     compiled[sheet_name][team].items(), key=lambda item: item[0]
#                 )
#             }
#             compiled[sheet_name][team] = {
#                 k: v
#                 for k, v in sorted(
#                     compiled[sheet_name][team].items(), key=lambda item: item[1]["Sub"]
#                 )
#             }
#     return compiled


def extract(
    data: dict[str, pd.DataFrame], replacements: dict[str, str]
) -> dict[str, dict[str, dict[str, dict[str, int]]]]:
    compiled = dict()
    for sheet_name, sheet in data.items():
        current_team = str()
        for row in sheet.itertuples():
            pl = str(row.Player).strip()
            if pl in replacements.keys():
                pl = replacements[pl]
            if pl.isupper():
                current_team = pl
                continue
            if current_team not in compiled.keys():
                compiled[current_team] = dict()
            if pl not in compiled[current_team].keys():
                compiled[current_team][pl] = dict()
            if sheet_name not in compiled[current_team][pl].keys():
                compiled[current_team][pl][sheet_name] = dict()
            compiled[current_team][pl][sheet_name] = {
                "GC": row.GC,
                "GT": row.GT,
                "DP": row.DP,
                "Sub": bool(row.Sub),
            }
        for team in compiled.keys():
            compiled[team] = dict(
                sorted(compiled[team].items(), key=lambda item: item[0])
            )

            compiled[team] = dict(
                sorted(
                    compiled[team].items(),
                    key=lambda item: item[1][list(item[1].keys())[0]]["Sub"],
                )
            )
    return dict(sorted(compiled.items()))


# %%
replacements = {
    "Will Brusseu": "William Brousseau",
    "William Bruso": "William Brousseau",
    "Thomas Jennings": "TJ Jakubowski",
    "Nathanel Weniger": "Nathaniel Weniger",
    "Caleb cash": "Caleb Cash",
    "Saul": "Saul Streachek",
    "Christian O'hara": "Christian O'Hara",
    "Joseph ODonnell": "Joseph O'Donnell",
    "Grace ODonnell": "Grace O'Donnell",
    "Jayden Kass": "Jaydan Kass",
}
compiled_stats = extract(raw_sheets, replacements=replacements)

###   FOR DEBUGGING ONLY   ###

with open("stats.json", "w") as file:
    j.dump(compiled_stats, file)

###   FOR DEBUGGING ONLY   ###

# %%
# import xlsxwriter as xlsx

workbook = xlsx.Workbook("Stats Conglomerate.xlsx")
worksheet = workbook.add_worksheet("Stats")
worksheet.freeze_panes(2, 1)


def write(row, col, string, cell_format):
    return worksheet.write(row, col, string, workbook.add_format(cell_format))


basic = workbook.add_format({"font_name": "Libre Franklin"})
bold = workbook.add_format({"bold": True, "font_name": "Libre Franklin"})
centered = workbook.add_format({"align": "center", "font_name": "Libre Franklin"})
first_column = workbook.add_format({"right": 2, "font_name": "Libre Franklin"})
other_columns = workbook.add_format({"right": 4, "font_name": "Libre Franklin"})

worksheet.set_column(0, 50, cell_format=basic)
worksheet.set_column(0, 0, cell_format=first_column)

worksheet.write(0, 0, "SEASON 9: STATISTICS", bold)
worksheet.write(1, 0, "MARAUDERS ULTIMATE FRISBEE LEAGUE", bold)

for i, week in enumerate(raw_sheet_names):
    i += 1
    worksheet.merge_range(0, (3 * i) - 2, 0, (3 * i), week, centered)
    worksheet.write(1, (3 * i) - 2, "Goals Thrown")
    worksheet.write(1, (3 * i) - 1, "Goals Caught")
    worksheet.write(1, (3 * i), "Defensive Plays")
    worksheet.set_column((3 * i), (3 * i), cell_format=other_columns)

row = 2
sheet_stats = dict()


def cols():
    ltrs = [*"ABCDEFGHIJKLMNOPQRSTUVWXYZ"]
    i = 0
    while True:
        if (i + 1) > len(ltrs):
            yield ltrs[i // len(ltrs) - 1] + ltrs[i % len(ltrs)]
        else:
            yield ltrs[i]
        i += 1


for team, team_stats in compiled_stats.items():
    col = cols()
    next(col)
    worksheet.write(row, 0, team, bold)
    for i in range(len(raw_sheet_names) * 3):
        c = next(col)
        worksheet.write(
            row, i + 1, f"=SUM({c}{row+2}:{c}{row + 1 + len(team_stats)})", bold
        )
    row += 1
    for player, player_stats in team_stats.items():
        worksheet.write(row, 0, player)
        for i, (week, week_stats) in enumerate(player_stats.items()):
            worksheet.write(row, 1 + (3 * i), week_stats["GT"])
            worksheet.write(row, 2 + (3 * i), week_stats["GC"])
            worksheet.write(row, 3 + (3 * i), week_stats["DP"])
        row += 1


worksheet.autofit()
workbook.close()
