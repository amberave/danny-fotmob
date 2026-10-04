# %%
# imports
import pandas as pd
import json
from bs4 import BeautifulSoup
import os
from openpyxl import load_workbook

# %%
# define league and year
LEAGUE = "FIFA World Cup"
YEAR = "2026/2027"

# %%
# base data
df = pd.read_excel("players.xlsx", index_col=False)
df

# %%
df[["https", "null", "fotmob", "players", "player_id", "slug"]] = df["FotMob URL"].str.split("/", expand=True)
df.drop(columns=["https", "null", "fotmob", "players"], inplace=True)
df

# %%
# get html data
html_files = os.listdir("html")
html_files[0]

# %%
# get the season id for that player
season_ids_by_player = {}
def get_season_id(filename):
    player_name = filename.split(' - ')[0]
    with open(f"html/{filename}", 'r', encoding='utf-8') as file:
        html_content = file.read()
    soup = BeautifulSoup(html_content, 'html.parser')
    div = soup.select("select")[1]
    try:
        optgroup = div.find("optgroup", {"label": YEAR})
        if optgroup is None:
            optgroup = div.find("optgroup", {"label": YEAR.split('/')[0]})
    
        option_tags = optgroup.select('select option')
        for option in option_tags:
            text = option.text
            if text == LEAGUE:
                if player_name == "Vítinha":
                    player_name = "Vitinha"
                season_ids_by_player[player_name] = option["value"]
    
    except Exception as e:
        print(f"Error: {player_name} did not have an entry for {YEAR} {LEAGUE} - {e}")

for file in html_files:
    get_season_id(file)
season_ids_by_player



# %%
df2 = df
df2['year'] = YEAR
df2['league'] = LEAGUE
df2['season_id'] = df2['Player'].map(season_ids_by_player)
# df.dropna(axis=0, how='any',subset=["player_id"], inplace=True)
df2

# %%
# get api endpoint
template = '"https://www.fotmob.com/api/data/playerStats?playerId="+df2[f"{player_id}"]+"&seasonId="+df2[f"{season_id}"]+"&isFirstSeason=false"'

player_id = "player_id"
season_id = "season_id"
df2["api_endpoint"] = eval(template)
df2

# %%
# export
league_and_year=f"{YEAR.replace('/', '-')} {LEAGUE}"
df2.to_excel(f"{league_and_year}.xlsx", index=False, sheet_name=league_and_year)
if not os.path.isdir(f"json/{league_and_year}"):
    os.mkdir(f"json/{league_and_year}")

