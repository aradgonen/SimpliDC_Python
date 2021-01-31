from openpyxl import load_workbook
from py2neo import Graph
db = Graph("bolt://neo4j:shMador1@localhost:7687")
def addnewrackstodb(db,data):
    data_rows = data.max_row

    for row in range(2, data_rows + 1):
        cur_rack = {}
        cur_node = {}
        cur_link = {}
        cur_rack['rack_id'] = data.cell(row, 1).value
        cur_rack['network'] = data.cell(row, 21).value
        cur_rack['size'] = 42
        create_racknode(db,cur_rack)
        if(data.cell(row,4).value == "Storage"):

            cur_node['arrayType'] = data.cell(row,5).value + " " + data.cell(row,7).value
            cur_node['serialNumber'] = "'"+data.cell(row,6).value+"'"
            cur_node['osVersion'] = data.cell(row,19).value
            cur_node['vendor'] = data.cell(row,5).value
            cur_node['extramgmtIps'] = data.cell(row,9).value
            cur_node['clusterName'] = data.cell(row,5).value
            cur_node['osType'] = data.cell(row,19).value
            cur_node['uNumber'] = data.cell(row,2).value
            cur_node['arrayProtocol'] = data.cell(row,4).value
            cur_node['name'] = data.cell(row,5).value
            cur_node['rackNumber'] = data.cell(row,1).value

            create_storagenode(db,cur_node)


def create_racknode(db, rackdata):
    rack_id = rackdata['rack_id']
    size = rackdata['size']
    network = rackdata['network']
    db.run('CREATE(n:RackNode {name:$rack_id, networkId:$network,size:$size})',rack_id=rack_id,size=size,network=network)

def create_storagenode(db,storagedata):
    arrayType = storagedata['arrayType']
    serialNumber = storagedata['serialNumber']
    osVersion = storagedata['osVersion']
    vendor = storagedata['vendor']
    extramgmtIps = storagedata['extramgmtIps']
    clusterName = storagedata['clusterName']
    osType = storagedata['osType']
    uNumber = storagedata['uNumber']
    arrayProtocol = storagedata['arrayProtocol']
    name = storagedata['name']
    rackNumber = storagedata['rackNumber']
    db.run('CREATE(n:StorageNode:XDeviceNode {arrayType:$arrayType,serialNumber:$serialNumber,osVersion:$osVersion,vendor:$vendor,extramgmtIps:$extramgmtIps,clusterName:$clusterName,osType:$osType,uNumber:$uNumber,arrayProtocol:$arrayProtocol,name:$name,rackNumber:$rackNumber})', arrayType=arrayType,serialNumber=serialNumber,osVersion=osVersion,vendor=vendor,extramgmtIps=extramgmtIps,
    clusterName=clusterName,osType=osType,uNumber=uNumber,arrayProtocol=arrayProtocol,name=name,rackNumber=rackNumber)
    put_in_rack(db,serialNumber,rackNumber)
def create_networknode(db,netwrokdata):
    return

def create_servernode(db,serverdata):
    return

def create_xdevicenode(db,xdevicedata):
    return

def create_linkrelation(db,linkdata):
    return

def put_in_rack(db,serialNumber,rackNumber):
    db.run("MATCH (a:StorageNode),(b:RackNode) WHERE a.serialNumber = '$serialNumber' AND b.name = '$rackNumber' CREATE (a)-[r:IN]->(b) RETURN type(r)",serialNumber=serialNumber,rackNumber=rackNumber)
wb = load_workbook("SampleDCMap.xlsx")
data = wb.active
# data_rows = data.max_row
# data_cols = data.max_column
# for col in range (1,data_cols):
#     for row in range(1,data_rows+1):
#         print(data.cell(row,col).value)
addnewrackstodb(db,data)



