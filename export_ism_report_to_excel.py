# This script receives as input a pdf file which represents the ISM report.
# It parses the PDF file and creates a excel file out of it. The excel file
# contains one sheet for each indicator available in the ISM report.
# And each sheet contains information for each industry for that specific
# indicator: if an industry reported growth, neutrality or decrease.
# Usage: python export_ism_report_to_excel.py input_file_name.pdf

from collections import OrderedDict
import os
from time import sleep

import fitz
import xlsxwriter

# Below is a list of all the possible industries in the ISM report
LIST_OF_INDUSTRIES = ["Textile Mills",
                      "Primary Metals",
                      "Transportation Equipment",
                      "Apparel, Leather & Allied Products",
                      "Petroleum & Coal Products",
                      "Printing & Related Support Activities",
                      "Machinery",
                      "Computer & Electronic Products",
                      "Miscellaneous Manufacturing",
                      "Electrical Equipment, Appliances & Components",
                      "Plastics & Rubber Products",
                      "Paper Products",
                      "Furniture & Related Products",
                      "Chemical Products",
                      "Food, Beverage & Tobacco Products",
                      "Fabricated Metal Products",
                      "Nonmetallic Mineral Products",
                      "Wood Products"]

# Below is a list containing all the indicators available in the ISM report
# Note: the PDF must contain all these indicators in this order, otheriwse it will break
LIST_OF_INDICATORS = ["NEW ORDERS",
                      "PRODUCTION",
                      "EMPLOYMENT",
                      "SUPPLIER DELIVERIES",
                      "INVENTORIES\n",
                      "CUSTOMERS' INVENTORIES",
                      "PRICES",
                      "BACKLOG OF ORDERS",
                      "NEW EXPORT ORDERS",
                      "IMPORTS",
                      "BUYING POLICY"]

def process_paragraph(para_text):
    """
    Process one text paragraph from the report which contains info about the
    industries which reported growth and decrease.
    This will return a dictionary containing key value pairs, where the key is the
    industry, and the value is the index.
    """
    sentences = para_text.split(".")
    growth_sentence = sentences[0]
    decrease_sentence = sentences[1]

    list_industries_growth = get_list_of_industries_from_sentence(growth_sentence)
    list_industries_decrease = get_list_of_industries_from_sentence(decrease_sentence)

    # find out the list of neutral industries
    list_industries_neutral = [industry for industry in LIST_OF_INDUSTRIES if industry not in list_industries_growth and industry not in list_industries_decrease]

    industries_dict = create_dict_of_industries(list_industries_growth, list_industries_neutral, list_industries_decrease)

    return industries_dict

def get_list_of_industries_from_sentence(sentence):
    """
    Creates a list of industries based on the given sentence. It parses the sentence and
    returns an ordered list of industries based on how they appear in the sentence.
    """
    list_industries = []

    if ":" in sentence:
        industries = sentence.split(":")[1]
        if ";" in industries:
            list_industries = [industry.replace(" and ", "").replace("\n", " ").strip() for industry in industries.split(";")]
        else:
            # only one industry after :
            list_industries.append(industries.replace("\n", " ").strip())
    else:
        # this is a special case in case there is no :
        for industry in LIST_OF_INDUSTRIES:
            if industry in sentence:
                list_industries.append(industry)

    return list_industries

def create_dict_of_industries(list_growth, list_neutral, list_decrease):
    """
    Creates a dictionary of key value pairs where the key is the industry,
    and the value is the index of growth/neutral/decrease.
    """
    industries_dict = OrderedDict()

    growth_index = len(list_growth)
    for growth_industry in list_growth:
        industries_dict[growth_industry] = growth_index
        growth_index -= 1

    for neutral_industry in list_neutral:
        industries_dict[neutral_industry] = 0

    decrease_index = -1
    for decrease_industry in list_decrease:
        industries_dict[decrease_industry] = decrease_index
        decrease_index -= 1

    return industries_dict

def export_dict_to_excel(industries_dict, output_filename):
    """
    Exports a given dictionary to excel.
    """
    workbook = xlsxwriter.Workbook(output_filename)
    for indicator, value in industries_dict.items():
        worksheet = workbook.add_worksheet(indicator)

        row = 0
        col = 0

        for industry, index in value.items():
            worksheet.write(row, col, industry)
            
            state = ""
            if index > 0:
                state = "Growth"
            elif index == 0 :
                state = "Neutral"
            else:
                state = "Contraction"

            worksheet.write(row, col + 1, state)
            worksheet.write(row, col + 2, index)

            row += 1

    workbook.close()

if __name__ == "__main__":
    
    filename = input("Enter the ISM report filename (ex: ism-report-2023.pdf): ")

    if not os.path.exists(filename):
        raise Exception("The provided file name does not exist")

    print("Opening input file and parsing the text...")
    with fitz.open(filename) as doc:
        text = ""
        for page in doc:
            text += page.get_text()

    industries_dict = OrderedDict()

    print("Parsing more the text inside the PDF to crate a list of indicators and industries...")
    previous_text = text.split(LIST_OF_INDICATORS[0])
    for i in range(1, len(LIST_OF_INDICATORS)):
        splitted_text = previous_text[1].split(LIST_OF_INDICATORS[i])
        if i == (len(LIST_OF_INDICATORS) - 1):
            paragraph = splitted_text[0].split(".\n")[-3]
        else:
            paragraph = splitted_text[0].split(".\n")[-2]
        
        previous_text = splitted_text

        processed_paragraph = process_paragraph(paragraph)
        industries_dict[LIST_OF_INDICATORS[i-1]] = processed_paragraph

    print("Generating XML file...")
    output_filename = filename.replace("pdf", "xlsx")
    export_dict_to_excel(industries_dict, output_filename)

    print("Excel was generated successfully!")
    print("Command prompt will close in 5 seconds! Beware!")
    sleep(5)
