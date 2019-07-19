import xlsxwriter
import xml.etree.ElementTree as ET

col = {}
row = {}

def add_children_to_sheet(parent):

    # If the worksheet has already been opened
    if parent.tag in col:
        col[parent.tag] = 0
        row[parent.tag] += 1
        c = col[parent.tag]
        r = row[parent.tag]

    # If the worksheet has not been opened yet
    else:
        col[parent.tag] = 0
        row[parent.tag] = 1
        c = col[parent.tag]
        r = row[parent.tag]

    for child in parent:
        c = col[parent.tag]
        col[parent.tag] += 1
        # If element has no children
        if len(list(child)) == 0:
            worksheet = workbook.get_worksheet_by_name(parent.tag)
            worksheet.write(0, c, child.tag)
            worksheet.write(r, c, child.text)
            print("R", r, "C", c)

        # If element does have children
        else:
            #If the worksheet has already been opened
            if child.tag in col:
                col[child.tag] += 1

            #If the worksheet worksheet doesn't exist
            if workbook.get_worksheet_by_name(child.tag) is None:
                workbook.add_worksheet(child.tag)
                row[child.tag] = 1
                col[child.tag] = 0

            add_children_to_sheet(child)

tree = ET.parse("BE Internal Auditing.xml")
#tree = ET.parse("input.xml")
root = tree.getroot()

workbook = xlsxwriter.Workbook("test.xlsx")
worksheet = workbook.add_worksheet(root.tag)

add_children_to_sheet(root)
workbook.close()