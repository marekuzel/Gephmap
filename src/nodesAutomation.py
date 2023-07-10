from openpyxl import Workbook, load_workbook

name_of_file = "export_emigrants_seniors_cz"

wb = load_workbook(name_of_file + ".xlsx")
ws = wb.active

#node_dict_appearances is used to store amount of unique nodes
#node_dict_totInt stores the number of interactions each unique node recieved
node_names = []
node_dict_appearances = {}
node_dict_totInt = {}

#for files in os.listdir():
#    if ".xlsx" in files:

#cycle for counting number of appearances for Groups, in excel this is done by COUNTIF
for cell in ws["A:A"]:
    if cell.value != "Page Name":
        if cell.value not in node_names:
            node_names.append(cell.value)
        if cell.value not in node_dict_appearances:
            node_dict_appearances[cell.value] = 1
        else:
            node_dict_appearances[cell.value] += 1

#total appearances for "Groups"
for node in node_dict_appearances:
    for i, cell in enumerate(ws["A:A"]):
        i += 1
        if node not in node_dict_totInt:
            node_dict_totInt[node] = 0
        if (node == cell.value and i>1):#i must be highere then 1 because openpyxl cant handle cell value of 0
            node_dict_totInt[node] += int(ws.cell(row=i, column=14).value) #column N always contains number of interaction for inidividual node



link_names = []
link_dict_appearances = {}
link_dict_totInt = {}           
 
#total interactions for "Links"
for cell in ws["AF:AF"]:
    if cell.value != "Link" or cell.value != "None":
        if cell.value not in link_names:
            link_names.append(cell.value)
    if cell.value not in link_dict_appearances:
        link_dict_appearances[cell.value] = 1
    else:
        link_dict_appearances[cell.value] += 1

#total appearances for "Links"
for link in link_dict_appearances:
    for i, cell in enumerate(ws["AF:AF"]):
        i += 1
        if link not in link_dict_totInt:
            link_dict_totInt[link] = 0
        if (link == cell.value and i>1):#i must be highere then 1 because openpyxl cant handle cell value of 0
            link_dict_totInt[link] += int(ws.cell(row=i, column=14).value) #column N always contains number of interaction for inidividual link



with open(name_of_file + "_Nodes.csv", 'w', encoding="utf-8") as f:
    f.write ("Label, Id, Appearances, Total Interactions, Type\n")
    for row in node_names:
        # Write the values in a row.
        f.write('"{}"'.format(row))
        f.write(",")
        f.write('"{}"'.format(row))
        f.write(",")
        f.write(str(node_dict_appearances[row]))
        f.write(",")
        f.write(str(node_dict_totInt[row]))
        f.write(',Group\n') # Add a new line
    for row in link_names:
        f.write('"{}"'.format(row))
        f.write(",")
        f.write('"{}"'.format(row))
        f.write(",")
        f.write(str(link_dict_appearances[row]))
        f.write(",")
        f.write(str(link_dict_totInt[row]))
        f.write(',Link\n') # Add a new line

with open(name_of_file + "_Edges.csv", 'w', encoding="utf-8") as f:
    f.write ("Source, Target\n")
    i=2
    for cell in ws["A:A"]:
        f.write('"{}"'.format(ws.cell(row=i, column=1).value))
        f.write(",")
        f.write('"{}"'.format(ws.cell(row=i, column=32).value))
        i+=1
        f.write("\n")