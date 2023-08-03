import sqlite3

agents = []
database = []
key_dict = {}

try:
    conn = sqlite3.connect('rcrg.db')
    c = conn.cursor()
    print("Successfully Connected to Database!")

    c.execute('SELECT * FROM agents')

    for row in c:
        agents.append(row)

except:
    print("Hello Error!")

finally:
    c.close()
    conn.close()
    print("Connection to Database Closed!")

i = 0

for agent in agents:
    
    database.append(agents[i][1] + " " + agents[i][2])
    
    i+=1



print(agents)
print(database)