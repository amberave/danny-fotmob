# FotMob data

*IMPORTANT*: You cannot scrape FotMob, it has specified on its site that you are not allowed to scrape its contents. When using these scripts, *you* become the scraper and the Python does the processing.

## General Overview

There are two scripts, [get_api_endpoints.py](./get_api_endpoints.py) and [parse_player_jsons_to_excel.py](./parse_player_jsons_to_excel.py).

The Excel file `[players.xlsx](./players.xlsx)` has the general player metadata needed for FotMob - DO NOT DELETE.

### Instructions
1. Decide which league and year you want and change the variables `LEAGUE` and `YEAR` at the top of **both** [get_api_endpoints.py](./get_api_endpoints.py) and [parse_player_jsons_to_excel.py](./parse_player_jsons_to_excel.py).

2. Run [get_api_endpoints.py](./get_api_endpoints.py) (double-click the file).

3. A new Excel file called "YEAR LEAGUE.xlsx" will appear. Open it.

4. In the column `api_endpoints`, there should be links. Copy the `player_id` from that row, click the link, Ctrl+S to save the page as a JSON file.

5. Paste the `player_id` as the filename, and save into [json/YEAR LEAGUE](./json) (same name as the file you have open). Continue Steps 4-5 for all players/links available.

6. **Close the Excel file** and run [parse_player_jsons_to_excel.py](./parse_player_jsons_to_excel.py). This will pull the data from the JSONs you just downloaded and put into the Excel file. 

7. Once the [Python script](./parse_player_jsons_to_excel.py) has finished running, open the Excel file again. It should now be filled with all the data from FotMob. Hooray!

