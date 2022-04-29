import pandas as pd
import random

# Read Excel Data
excel_data_df = pd.read_excel(r'C:\Users\bpham\pythonProject\Micron21 Audit 2022(1).xlsx', sheet_name='KVM')

# Read Company Name Column
Company_Name = (excel_data_df['KVM / Micron21 KVM'].tolist())
Company_Name.pop(0)

# Read Client ID Column
Client_ID = (excel_data_df['Unnamed: 1'].tolist())
Client_ID.pop(0)

# Read Name Column
Name = (excel_data_df['Unnamed: 2'].tolist())
Name.pop(0)

# Read IP Address Column
IP_Address = (excel_data_df['Unnamed: 3'].tolist())
IP_Address.pop(0)

id_array = [] # id
id2_array = [] # typeid / parentid

# Generates 12 character num
def random_num():
    return random.randint(100000000000, 999999999999)

# export xml format
def xml_export():
    array_size = len(Company_Name)
    id_array.append(random_num())
    id2_array.append(random_num())
    typeid = 16

    array_length = 0

    print('<export>')

    for i in range(array_size):
        if str(Company_Name[i]) == "nan" or str(Client_ID[i]) == "nan" or str(Name[i]) == "nan" or str(IP_Address[i]) == "nan":
            break
        else:
            id_array.append(random_num())
            id2_array.append(random_num())
            print('<monitor>')
            print(f'<id>{id_array[i]}</id>')
            print(f'<typeid>{typeid}</typeid>')
            print("<nv>")
            print("<ui>")
            print("<name>PING</name>")
            print("<server></server>")
            print("</ui>")
            print("<cred/>")
            print(f"<addr>{IP_Address[i]}</addr>")
            print("</nv>")
            print("<testing>")
            print("<up>20</up>")
            print("<warn>20</warn>")
            print("<down>5</down>")
            print("<lost>5</lost>")
            print("</testing>")
            print("<stats>")
            print("<enabled>true</enabled>")
            print("</stats>")
            print(f"<parentid>{id2_array[i]}</parentid>")
            print("<enabled>true</enabled>")
            print("<maxtest>20</maxtest>")
            print("<notifyfailures>2</notifyfailures>")
            print("<maxalerts>1</maxalerts>")
            print("<tags/>")
            print("</monitor>")
            array_length += 1

    ## KVM Customers Group ##
    print("<group>")
    print("<id>418261689554</id>")
    print("<typeid>3899997</typeid>")
    print("<parentid>23124513034</parentid>")
    print("<nv>")
    print("<ui>")
    print("<name>KVM-Customers</name>")
    print("</ui>")
    print("</nv>")
    print("<col>")
    for i in range(array_length):
        print(f'<u{i}>{id2_array[i]}</u{i}>')
    print("</col>")
    print("<depends/>")
    print("</group>")
    ##########################



    ## Device Information & Tag ##
    for i in range(array_length):
        print("<group>")
        print(f"<id>{id2_array[i]}</id>")
        print(f"<typeid>3899973</typeid>")
        print("<nv>")
        print("<ui>")
        print(f"<name>{Client_ID[i]} - {Company_Name[i]}</name>")
        print("</ui>")
        print("</nv>")
        print("<parentid>0</parentid>")
        print("<propsro>false</propsro>")
        print("<tags>")
        print("<tag>")
        print("<name>IP</name>")
        print(f"<value>{IP_Address[i]}</value>")
        print("</tag>")
        print("<tag>")
        print("<name>ID</name>")
        print(f"<value>{Client_ID[i]}</value>")
        print("</tag>")
        print("<tag>")
        print("<name>Name</name>")
        print(f"<value>{Name[i]}</value>")
        print("</tag>")
        print("</tags>")
        print("<col>")
        print(f"<u{i}>{id_array[i]}</u{i}>")
        print("</col>")
        print("<depends/>")
        print("</group>")
    ###############################

    # Ending export tag
    print("</export>")




xml_export()
