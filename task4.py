import sqlite3
from tokenize import Name
import pandas as pd
from sqlalchemy import create_engine

file = "C:\\Users\\Malik\\OneDrive\\Desktop\\sample.xlsx"
# output = "output.xlsx"
engine = create_engine('sqlite://', echo=False)
df = pd.read_excel(file, sheet_name="Sheet1")
df.to_sql('sample', engine, if_exists = 'replace', index=False)
# result = engine.execute("SELECT Email, Max(Sale) FROM sample where Age >=20 AND Age <=25")
result = engine.execute(" Select Email from sample where Sale = (Select Max(Sale) from sample)")
# result = engine.execute("Select Email from sample where Sale = 'hIGHEST'")

email_25000 = pd.DataFrame(result)                          
# final.to_excel(output, index=False)
print(email_25000)