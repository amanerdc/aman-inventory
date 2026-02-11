# Aman Inventory

A simple desktop inventory app for Unica and HDN Integrated Farm.

## Run
```powershell
python app.py
```

## Default Login
- Username: `admin`
- Password: `admin123`

## Notes
- Data is stored locally in PostgreSQL.
- Unica Perishable tracks IN/OUT and stock balances.
- HDN Plants is a placeholder screen for future farm inventory.

## PostgreSQL Setup
Create a local database (example name `aman_inventory`) and then add a `db_config.json` file in this folder
(use `db_config.json.example` as a starting point):
```json
{
  "host": "localhost",
  "port": 5432,
  "dbname": "aman_inventory",
  "user": "postgres",
  "password": "postgres"
}
```

You can also set a DSN with `AMAN_DB_URL`.

## Optional Export Dependencies
Exports will still work with basic fallbacks, but for best results install:
```powershell
pip install openpyxl reportlab pillow
```

## UI Extras
Calendar pop-outs and image preview use:
```powershell
pip install tkcalendar pillow
```
