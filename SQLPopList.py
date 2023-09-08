import sqlite3

def SQLPopList(table_name, list_name):
    
    new_list = []

    try:
        conn = sqlite3.connect('rcrgbrokerage.db')
        c = conn.cursor()
        print("Successfully Connected to Database!")

        c.execute(f'SELECT * FROM {table_name}')

        for row in c:
                new_list.append(row)

    except:
        print("Hello Error!")

    finally:
        c.close()
        conn.close()
        print("Connection to Database Closed!")

    list_name = new_list
