import sqlite3

def SQLPopList(table):
    
    lst = []
    database = []
    key_dict = {}

    try:
        conn = sqlite3.connect('rcrg.db')
        c = conn.cursor()
        print("Successfully Connected to Database!")

        c.execute(f'SELECT * FROM {table}')

        for row in c:
            lst.append(row)

    except:
        print("Hello Error!")

    finally:
        c.close()
        conn.close()
        print("Connection to Database Closed!")

    i = 0
    for row in lst:
        database.append(lst[i][1] + " " + lst[i][2])
        i+=1

    j =0
    for row in lst:
        key_dict[database[j]] = row
        j += 1

    print(key_dict)

SQLPopList("agents")
