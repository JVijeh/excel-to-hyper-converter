#  Excel‑to‑Hyper Converter: Solar Energy Example

This project demonstrates how to convert an Excel file (`RWFD_Solar_Energy.xlsx`) into Tableau `.hyper` extract files using Python. It’s designed for data analysts—especially those familiar with Tableau—who are learning or revisiting Python.

You’ll explore two methods:
- **Method 1 – Pantab Express**: A quick, one-liner conversion using the Pantab library.
- **Method 2 – Hyper API Adventure**: A more hands-on approach using Tableau’s official Hyper API for full control over schema and data insertion.

---

##  Requirements

Install the required Python packages:

```bash
pip install pandas pantab tableauhyperapi
```

## Project Structure
project-root/
│
├── src/
│   ├── method1_pantab.py       # Pantab method (fast and simple)
│   └── method2_hyperapi.py     # Hyper API method (hands-on control)
│
├── video/                      # Demo videos (optional)
├── .gitignore
└── README.md

## Method 1 – Pantab Express 
This method uses the Pantab library to convert a DataFrame to a .hyper file.

```bash
import pandas as pd
import pantab
```

# Load Excel data into a DataFrame
```bash
df = pd.read_excel(
    r"C:\\Users\\jvije\\OneDrive\\DataDevQuest\\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)
```
# Convert to .hyper file
```bash
pantab.frame_to_hyper(df,
    r"C:\\Path\\To\\Output\\RWFD_Solar_Energy_Method1_Pantab.hyper",
    table="Extract")
```

# Run it with:
```bash
python src/method1_pantab.py
```

## Method 2 – Hyper API Adventure
This method gives you full control over the schema and data insertion process using Tableau’s Hyper API.

```bash
import pandas as pd
from tableauhyperapi import HyperProcess, Connection, TableDefinition, SqlType, TableName, Inserter, CreateMode, Nullability
```

# Load Excel data into a DataFrame
```bash
df = pd.read_excel(
    r"C:\\Users\\jvije\\OneDrive\\DataDevQuest\\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)
```
# Start the Hyper process
```bash
with HyperProcess() as hyper:
    with Connection(endpoint=hyper.endpoint,
                    database=r"C:\\Path\\To\\Output\\RWFD_Solar_Energy_Method2_HyperAPI.hyper",
                    create_mode=CreateMode.CREATE_AND_REPLACE) as conn:

        conn.catalog.create_schema("Extract")

        table_def = TableDefinition(
            table_name=TableName("Extract", "SolarData"),
            columns=[TableDefinition.Column(col, SqlType.text(), Nullability.NULLABLE) for col in df.columns]
        )

        conn.catalog.create_table(table_def)

        with Inserter(conn, table_def) as inserter:
            inserter.add_rows(rows=df.values.tolist())
            inserter.execute()
```

# Run it with:
```bash
python src/method2_hyperapi.py
```

## Demo


## What to Expect
- **Two .hyper files generated (one for each method).
- **You should be able to open the Tableau Extracts using Tableau Desktop to verify the data.
- **You should also receive an output that shows how long the entire process took in order to compare.


