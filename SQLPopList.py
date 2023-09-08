import sqlite3

def SQLPopList(table_name):
    
    new_list = []
    key_database = []
    dict = {}

    try:
        conn = sqlite3.connect('rcrgbroker.db')
        c = conn.cursor()
        print("Successfully Connected to Database!")

        c.execute(f'SELECT * FROM {table_name}')

        for row in c:
            new_list.append(row)

    except:
        print("There was an error connecting to the RCRG Database.")

    finally:
        c.close()
        conn.close()
        print("Connection to Database Closed.")

    
    if table_name == 'rcrg':
        for i, row in enumerate(new_list):
            key_database.append(new_list[i][1] + " " + new_list[i][2])
    else:
        for i, row in enumerate(new_list):
            key_database.append(new_list[i][1] + " (" + new_list[i][2] + ")")


    for j, entry in enumerate(new_list):
            dict[key_database[j]] = entry

    return key_database, dict