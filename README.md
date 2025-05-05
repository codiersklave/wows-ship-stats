# World of Warships PlayerStats

## Download Ship Info from API (with Images)
```bash
python ships.py
```

## Create Ship ID List
```bash
python ids.py
```

## Download Player Stats
```bash
python stats.py --help

----------
usage: stats.py [-h] [--days DAYS] [--no-description] [--type {all,A,B,C,D,S}] [--nation {all,CW,EU,FR,DE,IT,JP,NL,AM,AS,ES,UK,US,SU}] [--order {date,name}] [--ship SHIP] account_id

Generate a docx file with stats for a given account.

positional arguments:
  account_id            Account ID

options:
  -h, --help            show this help message and exit
  --days DAYS           Number of days to include in the docx file.
  --no-description      Do not include ship descriptions in the docx file.
  --type {all,A,B,C,D,S}
                        Filter by ship type.
  --nation {all,CW,EU,FR,DE,IT,JP,NL,AM,AS,ES,UK,US,SU}
                        Filter by nation.
  --order {date,name}   Order by date or name.
  --ship SHIP           The ship ID we are interested in. If specified, only this ship will be included in the docx file.
```
