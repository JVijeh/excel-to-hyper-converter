import pandas as pd
from tableauhyperapi import HyperProcess, Connection, TableDefinition, SqlType, TableName, Inserter, CreateMode, Nullability

# Load Excel data into a DataFrame
df = pd.read_excel(
    r"C:\\Users\\jvije\\OneDrive\\DataDevQuest\\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)

# Start the Hyper process
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
