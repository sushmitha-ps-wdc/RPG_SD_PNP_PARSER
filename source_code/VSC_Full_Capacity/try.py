import openpyxl
from xml.etree import ElementTree as ET


def get_chart_info(excel_file):
    try:
        workbook = openpyxl.load_workbook(excel_file)

        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]

            # Check for charts in the sheet
            if sheet._charts:
                print(f"Sheet: {sheet_name}")

                # Initialize a counter for the number of charts
                chart_count = 0

                # Iterate through each chart in the sheet
                for chart in sheet._charts:
                    # Get the chart title from XML or 'No Title' if not present
                    chart_title = chart.title.text if chart.title else 'No Title'
                    if chart_title == 'No Title':
                        # If title is 'No Title', try to extract it from XML
                        chart_title = extract_chart_title(chart)

                    print(f"  Chart {chart_count + 1}: {chart_title}")

                    # Increment the chart count
                    chart_count += 1

                print(f"  Total Charts: {chart_count}\n")

    except Exception as err:
        print(f"Error occurred: {err}")


def extract_chart_title(chart):
    xml_string = chart._chartSpace
    root = ET.fromstring(xml_string)

    title_element = root.find('.//c:title', namespaces={'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'})

    if title_element is not None:
        return title_element.text

    return 'No Title'

excel_file_path = r"C:\Users\42395\Desktop\Automation\Parser_Creation\UHS-II\VSC Full Capacity\29\VSC_Full_Capacity.xlsx"
get_chart_info(excel_file_path)
