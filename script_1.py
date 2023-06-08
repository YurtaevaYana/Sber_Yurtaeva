import pandas as pd
from lxml import etree as et

data = pd.read_excel(r'test_input.xlsx', skiprows = 4, dtype='object')
filename = (pd.read_excel(r'test_input.xlsx', header = 0, skiprows = 1, nrows = 1, usecols = 'B').iloc[0])[0]
data['IE Code'] = data['IE Code'].apply(lambda x: '{0:0>10}'.format(x))
data['Issuance Date'] = pd.to_datetime(data['Issuance Date']).dt.date
data['SB Date'] = pd.to_datetime(data['SB Date']).dt.date
data['Client'] = data['Client'].apply(lambda i: f'"{i}"')

root = et.Element('CERTDATA')
root_heading = et.SubElement(root, 'FILENAME')
root_heading.text = filename
sub_root = et.SubElement(root, 'ENVELOPE')
for row in data.iterrows():
    root_tags = et.SubElement(sub_root, 'ECERT')
    column_heading_1 = et.SubElement(root_tags, 'CERTNO')
    column_heading_2 = et.SubElement(root_tags, 'CERTDATE')
    column_heading_3 = et.SubElement(root_tags, 'STATUS')
    column_heading_4 = et.SubElement(root_tags, 'IEC')
    column_heading_5 = et.SubElement(root_tags, 'EXPNAME')
    column_heading_6 = et.SubElement(root_tags, 'BILLID')
    column_heading_7 = et.SubElement(root_tags, 'SDATE')
    column_heading_8 = et.SubElement(root_tags, 'SCC')
    column_heading_9 = et.SubElement(root_tags, 'SVALUE')
    
    column_heading_1.text = str(row[1]['Ref no'])
    column_heading_2.text = str(row[1]['Issuance Date'])
    column_heading_3.text = str(row[1]['Status'])
    column_heading_4.text = str(row[1]['IE Code'])
    column_heading_5.text = str(row[1]['Client'])
    column_heading_6.text = str(row[1]['Bill Ref no'])
    column_heading_7.text = str(row[1]['SB Date'])
    column_heading_8.text = str(row[1]['SB Currency'])
    column_heading_9.text = str(row[1]['SB Amount'])

tree = et.ElementTree(root)
et.indent(tree, space="\t", level=0)
tree.write('output_1.xml', encoding="utf-8")