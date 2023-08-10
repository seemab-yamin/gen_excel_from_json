import pandas as pd
import json
import os

def gen_excel_from_json(file_path = "dream_result.json", excel_file_path = 'output_file.xlsx'):
    # json file path
    if not os.path.exists(file_path):
        print("%s JSON File not found in current directory."%file_path)
    
    try:
        # read json file
        with open(file_path, "r") as file:
            data_obj = json.load(file)
    except Exception as e:
        print("Error Reading JSON File:\t%s"%e)
        return False
    
    try:
        # load sheet 1 headers
        table_headers = data_obj.get("table_headers")

        # list of dictionaries
        savings_list = data_obj.get("savings_list")

        # create df1 object for sheet 1
        df1 = pd.DataFrame(savings_list)

        # sheet 2 data from json obj
        data_sheet2 = {
                    "Total Months" : [data_obj.get("total_months")],
                    "Total Years" : [data_obj.get("total_years")],
                    "Additional Months" : [data_obj.get("additional_months")]
                }
        
        # create df2 object for sheet 2
        df2 = pd.DataFrame(data_sheet2)
    except Exception as e:
        print("Error JSON structure is changed:\t%s"%e)
        return False

    try:
        excel_writer = pd.ExcelWriter(excel_file_path)
        
        # Write dataframes to different sheets
        df1.columns=table_headers
        df1.to_excel(excel_writer, sheet_name='Sheet1', index=False)
        df2.to_excel(excel_writer, sheet_name='Sheet2', index=False)

        # Save the Excel file
        excel_writer.close()
        
        return excel_file_path
    except Exception as e:
        print("Error Saving Excel File:\t%s"%e)
        return False
    
if __name__ == "__main__":

    excel_file_path = gen_excel_from_json()
    if excel_file_path:
        print("PATH:\t%s"%excel_file_path)
        print("Excel generated successfully.")
    else:
        print("Excel NOT generated.")