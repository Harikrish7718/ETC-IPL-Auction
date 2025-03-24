import pandas as pd
import re

file_path = "ipl.txt"  
with open(file_path, "r", encoding="utf-8") as file:
    lines = [line.strip() for line in file if line.strip()]  

cleaned_lines = []
start_processing = False  

for line in lines:
    if line.startswith("1"):  
        start_processing = True

    if start_processing:
        cleaned_lines.append(line)

player_points = {}
i = 0  

while i < len(lines) - 3:  
    if lines[i].isdigit():  
        player_name = lines[i + 1].strip()  
        team_name = lines[i + 2].strip()  
        stats_line = lines[i + 3].strip()  

        stats_parts = stats_line.split()
        if stats_parts and stats_parts[0].replace(".", "", 1).isdigit():  
            points = float(stats_parts[0])  
            player_points[player_name] = points
        else:
            print(f"Skipping invalid data for player: {player_name}")
    i += 1  

excel_file = "ETC IPL Auction '25  Mastersheet 22_Mar.xlsx"  
df = pd.read_excel(excel_file, sheet_name="Players")

df['Player Points'] = df['Player Name'].map(player_points)

df.to_excel('ETC IPL.xlsx', index=False)

print("Excel file updated successfully: ETC IPL Updated.xlsx")