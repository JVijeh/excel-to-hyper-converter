import pandas as pd
import pantab

# Load Excel data into a DataFrame
df = pd.read_excel(
    r"C:\\Users\\jvije\\OneDrive\\DataDevQuest\\RWFD_Solar_Energy.xlsx",
    sheet_name="Actuals"
)

# Convert to .hyper file
pantab.frame_to_hyper(df,
    r"C:\\Path\\To\\Output\\RWFD_Solar_Energy_Method1_Pantab.hyper",
    table="Extract")
