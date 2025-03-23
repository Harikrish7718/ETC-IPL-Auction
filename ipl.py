import pandas as pd

# Load your existing Excel sheet
excel_file = 'ETC IPL.xlsx'  # Replace with your file path
df = pd.read_excel(excel_file, sheet_name='Players') 

# Create a dictionary of player names and points
player_points = {
    "Sunil Narine": 31.5,
    "Phil Salt": 29.5,
    "Ajinkya Rahane": 29.0,
    "Josh Hazlewood": 23.0,
    "Virat Kohli": 20.5,
    "Krunal Pandya": 18.5,
    "Rajat Patidar": 16.0,
    "Rasikh Dar": 15.0,
    "Varun Chakaravarthy": 10.5,
    "Liam Livingstone": 10.5,
    "Yash Dayal": 10.5,
    "Jitesh Sharma": 10.0,
    "Spencer Johnson": 9.5,
    "Vaibhav Arora": 7.5,
    "Harshit Rana": 6.5,
    "Rinku Singh": 5.0,
    "Devdutt Padikkal": 2.5,
    "Venkatesh Iyer": 2.5,
    "Ramandeep Singh": 2.5,
    "Andre Russell": 2.5,
    "Quinton De Kock": 2.5
}

# Update points in the DataFrame
df['Player Points'] = df['Player Name'].map(player_points)

# Save the updated Excel file
df.to_excel('ETC IPL.xlsx', index=False)
