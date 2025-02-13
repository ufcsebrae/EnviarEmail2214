import sqlalchemy

def consulta_sql():
    servername = "spsvsql39\\metas"
    dbname = "FINANCA"
    engine = sqlalchemy.create_engine(
        f'mssql+pyodbc://@{servername}/{dbname}?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server'
    )
    
    try:
        with engine.connect() as connection:
            print("Hello World!")
    except Exception as e:
        print("erro")
        raise

    return 