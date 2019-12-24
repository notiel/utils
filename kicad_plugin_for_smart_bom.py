"""
    @package
    Command line:
    python "pathToFile/bom_csv_sorted_by_ref.py" "%I" "%O"
"""

from __future__ import print_function
import os
import sys
import xlsxwriter
import datetime

# Import the KiCad python helper module
sys.path.append("C:\\Program Files\\KiCad\\bin\\scripting\\plugins")
import kicad_netlist_reader

# Create output dir if not exists
dirname = os.path.join(os.path.split(sys.argv[1])[0], "Factory")
if not os.path.exists(dirname):
    os.mkdir(dirname)
    print("Dir created")

# Construct output filename and create workbook
date_str = a = datetime.datetime.today().strftime("%d-%m-%Y")
fname = os.path.join(dirname, "BOM_%s.xlsx" % date_str)
workbook = xlsxwriter.Workbook(fname)
worksheet = workbook.add_worksheet('Bill of Materials')

# Add formats of cells
hdr_format = workbook.add_format()
hdr_format.set_bg_color('#C0C0C0')
hdr_format.set_bottom(1)
hdr_format.set_left(1)
hdr_format.set_right(1)

all_format = workbook.add_format()
all_format.set_bottom(1)
all_format.set_left(1)
all_format.set_right(1)

# Adjust column widths
worksheet.set_column(0, 1, 18)
worksheet.set_column(2, 2, 26)
worksheet.set_column(3, 10, 19)
worksheet.set_column(11, 11, 8)

# Output a header line
header = ['Type', 'Value', 'PN', 'Manufacturer', 'PN Alternative 1', 'PN Alternative 2', 'Designator', 'Footprint',
          'Dielectric', 'Tolerance', 'Description', 'Quantity']
worksheet.write_row(0, 0, header, hdr_format)

# Generate an instance of a generic netlist, and load the netlist tree from
# the command line option. If the file doesn't exist, execution will stop
net = kicad_netlist_reader.netlist(sys.argv[1])
components = net.getInterestingComponents()
grouped = net.groupComponents(components)

SMART = True  # use this flag to use smart type matcher (otherwise type is taken from table below)

# list of correct types. Use lower case letters and multiple form. 'capacitors' is correct, 'Capacitor' has two mistakes
correct_types = ['capacitors', 'resistors', 'inductors', 'ic', 'diodes', 'leds', 'connectors', 'installations',
                 'antennas', 'pictures', 'btnsswitches', 'quartz']

# use this dict to match types with concrete pns. Use capital letters of partnumber, 'MSD3C031V', not 'msd3c031v'
type_matcher_by_pn = {'MSD3C031V': 'Bidir Zener', 'SY8120': 'IC', 'STM32F411CxU6': 'IC', 'ICN2012': 'IC',
                      'ICN2595': 'IC', 'AT24C01D': 'IC'}

# use this dict to type correction after type was matched. Be careful with capital and lower case letters
type_corrector = {'leds': 'LED RGB', 'capacitors': 'Capacitor SMD', 'resistors': 'Resistor SMD',
                  'inductors': 'Inductor SMD', 'ic': 'IC', 'connectors': 'Connector', 'transistors': 'Transistor'}

# use this dict to match type by designator (Default way if SMART set to False)
type_matcher_by_designator = {'C': 'Capacitor SMD', 'DA': 'IC', 'DD': 'IC', 'D': 'Diode', 'Hole': 'Do not mount',
                              'Logo': 'Do not mount', 'Q': 'Transistor', 'L': 'Inductor SMD', 'R': 'Resistor SMD',
                              'SW': 'Swith or button', 'TP': 'do not mount', 'XL': 'Connector', 'XTAL': 'Quartz'}

# usr this dict for ic and other specific description, only capacitors, inductors and resistors can be
# descripted automatically. You may also use this dict to change default description
description_dict = {'SY8120': 'Sync Power Supply', 'MSD3C031V': 'ESD Protection Zener',
                    'ICN2595': '16-ch LED current supply', 'AT24C01D': 'EEPROM 1k', 'Choke': 'Choke'}


def get_type(partnumber: str, footprint: str, designator: str) -> str:
    """
    gets type of compoment using matching dictionaries with footprint (if smart parameter is enabled) or designator.
    also it is possible to specify type to concrete partnumber
    :param partnumber: partnumber of component
    :param footprint: footprint of component
    :param designator: designator of component
    :return: type of component
    """
    res = ""
    if partnumber.upper() in type_matcher_by_pn.keys():
        res = type_matcher_by_pn[partnumber.upper()]
    elif SMART and ':' in footprint:
        if footprint.split(':')[0].lower() in correct_types:
            res = footprint.split(':')[0].lower()
        elif designator in type_matcher_by_designator.keys():
            res = type_matcher_by_designator[designator]
    elif designator in type_matcher_by_designator.keys():
        res = type_matcher_by_designator[designator]
    if res.lower() in type_corrector.keys():
        return type_corrector[res.lower()]
    return res


def get_isolator(value: str) -> str:
    """
    returns isolator for capacitors according to theis value
    :param value: capacitor value
    :return: capacitor isolator
    """
    if 'pf' in value.lower():
        return 'NP0'
    return 'x5r or x7r'


def get_tolerance(component_type: str) -> str:
    """
    returns tolerance 1% for resistors and 20% for other components
    :param component_type: type of component
    :return:
    """
    if 'resistor' in component_type.lower():
        return '1%'
    elif 'capacitor' in component_type.lower():
        return '20%'
    else:
        return '-'


def get_description(component_type: str, value: str, footprint: str) -> str:
    """
    makes description for component
    :param component_type: type of component
    :param value: component value
    :param footprint: component footprint
    :return: description string
    """
    if value.lower() in description_dict.keys():
        return description_dict[value.lower()]
    case = ''.join([char for char in footprint if char.isdigit()])
    template = "Any %s %s value" % (value, component_type)
    if 'capacitor' in component_type.lower():
        return template + " with %s isolator in %s case" % (get_isolator(value), case)
    if 'resistor' in component_type.lower():
        return template + ' in ' + case + r'case with 1% tolerance'
    if 'inductor' in component_type.lower():
        return template + 'in %s case' % case
    return ''


RowN = 1
for group in grouped:
    # Add the reference of every component in the group and keep a reference
    # to the component so that the other data can be filled in once per group

    refs = ', '.join([component.getRef() for component in group if not component.getField('DoNotBOM')])

    # refs = ""
    # for component in group:
    #     if len(refs) > 0:
    #         refs += ", "
    #     refs += component.getRef()
    #     c = component
    #  if c.getField('DoNotBOM'):
    #     continue
    if refs:
        c = group[0]
        footprint = c.getFootprint()
        pn = c.getField("PN")
        designator = refs.split(',')[0]
        designator_letters = "".join([char for char in designator if not char.isdigit()])
        component_type = c.getField("Type") if c.getField("Type") else get_type(pn, footprint, designator_letters)
        value = c.getValue()
        isolator = c.getField("Dielectric")
        if not isolator and 'capacitor' in component_type.lower():
            isolator = get_isolator(value)
        tolerance = c.getField("Tolerance") if c.getField("Tolerance") else get_tolerance(component_type)
        description = c.getField("Description") if c.getField("Description") else get_description(component_type,
                                                                                                  value, footprint)
        worksheet.write(RowN, 0, component_type, all_format)
        worksheet.write(RowN, 1, c.getValue(), all_format)
        worksheet.write(RowN, 2, pn, all_format)
        worksheet.write(RowN, 3, c.getField("Manufacturer"), all_format)
        worksheet.write(RowN, 4, c.getField("PN Alternative 1"), all_format)
        worksheet.write(RowN, 5, c.getField("PN Alternative 2"), all_format)
        worksheet.write(RowN, 6, refs, all_format)  # Designator
        worksheet.write(RowN, 7, footprint, all_format)
        worksheet.write(RowN, 8, isolator, all_format)
        worksheet.write(RowN, 9, tolerance, all_format)
        worksheet.write(RowN, 10, description, all_format)
        worksheet.write(RowN, 11, len(group), all_format)  # Quantity
        RowN += 1

# Output all of the component information (One component per row)
# for c in components:
# if not c.getField("DoNotBOM"):
#     writerow(out, [c.getRef(), c.getValue(), c.getFootprint(),
#              c.getField("PN"), c.getField("Comment")])

workbook.close()
