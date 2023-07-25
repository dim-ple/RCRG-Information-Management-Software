import sqlite3

conn = sqlite3.connect('rcrg.db')

#Create our Cursor
c = conn.cursor()

#Create our Table
c.execute("""CREATE TABLE agents (
        agentid INTEGER NOT NULL PRIMARY KEY,
        agentfirst TEXT NOT NULL,
        agentlast TEXT NOT NULL,
        agentphone TEXT NOT NULL,
        agentemail TEXT NOT NULL,
        agenttype TEXT NOT NULL,
        agentlicensenum TEXT NOT NULL,
        agentbroker TEXT NOT NULL                 
    )""")

c.execute("""CREATE TABLE lenders (
        lenderid INTEGER NOT NULL PRIMARY KEY,
        lendercompany TEXT NOT NULL,
        lofirst TEXT NOT NULL,
        lolast TEXT NOT NULL,
        lophone TEXT NOT NULL,
        loemail TEXT NOT NULL,
        lpemail TEXT,           
    )""")

c.execute("""CREATE TABLE clients (
        clientid INTEGER NOT NULL PRIMARY KEY,
        clientfirst TEXT NOT NULL,
        clientlast TEXT NOT NULL,
        clientphone TEXT NOT NULL,
        clientemail TEXT NOT NULL,
        mailingstreetnum TEXT,
        malingstreetname TEXT,
        malingstreettype TEXT,
        malingcity TEXT,
        malingstate TEXT,
        malingzip TEXT,
        FOREIGN KEY(agentid) REFERENCES agents(agentid),
        FOREIGN KEY(lenderid) REFERENCES lenders(lenderid)                  
    )""")

c.execute("""CREATE TABLE hoas (
        hoaid INTEGER NOT NULL PRIMARY KEY,
        hoaname TEXT NOT NULL,
        hoamgmtco TEXT NOT NULL,
        hoaphone TEXT,
        hoaemail TEXT,               
    )""")

c.execute("""CREATE TABLE properties (
        propid INTEGER NOT NULL PRIMARY KEY,
        propstreetnum TEXT NOT NULL,
        propstreetname TEXT NOT NULL,
        propstreettype TEXT NOT NULL,
        propcity TEXT NOT NULL,
        propstate TEXT NOT NULL,
        propzip TEXT NOT NULL,
        FOREIGN KEY(hoaid) REFERENCES hoas(hoaid)                 
    )""")

#Commit our command
conn.commit()

#Terminate our connection
conn.close()