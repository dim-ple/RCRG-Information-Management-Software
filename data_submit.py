
agent_cols = "(agentfirst, agentlast, agentphone, agentemail, agenttype, agentlicensenum, agentbroker)"

def data_submit(table_name, *cols):
    ins_statement = 'INSERT INTO '
    val_statement = 'VALUES ("'
    end_char = ')'
    list_string = []
    
    if table_name == 'agents':
        ins_statement += agent_cols

    for col in cols:
        val_statement += col + '", "'

    
    list_string = list(val_statement)
    list_string.pop(len(list_string)-1)
    list_string.pop(len(list_string)-1)
    list_string.pop(len(list_string)-1)
    new_string = "".join(list_string)

    val_statement = new_string + end_char

    print(ins_statement)
    print(val_statement)



data_submit('agents', 'Tunde', 'Lewis-Leffew', '804-307-7008', 'tundehasthekey@gmail.com', 'Salesperson', '022526578', 'The Rick Cox Realty Group')