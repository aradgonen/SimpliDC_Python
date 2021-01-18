from openpyxl import load_workbook
from py2neo import Graph
db = Graph("bolt://neo4j:shMador1@localhost:7687")
def addnewrackstodb(db,data):
    data_rows = data.max_row
    cur_rack = {}
    for row in range(2, data_rows + 1):
        cur_rack['rack_id'] = data.cell(row, 1).value
        cur_rack['network'] = data.cell(row, 21).value
        cur_rack['size'] = 42
        create_racknode(db,cur_rack)
def create_racknode(db, rackdata):
    rack_id = rackdata['rack_id']
    size = rackdata['size']
    network = rackdata['network']
    db.run('CREATE(n:RackNode {name:$rack_id, networkId:$network,size:$size})',rack_id=rack_id,size=size,network=network)


wb = load_workbook("SampleDCMap.xlsx")
data = wb.active
# data_rows = data.max_row
# data_cols = data.max_column
# for col in range (1,data_cols):
#     for row in range(1,data_rows+1):
#         print(data.cell(row,col).value)
addnewrackstodb(db,data)



