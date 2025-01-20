import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt
import os
import argparse

column_dict = {}

def sheet_generator(path_name, output_excel_file):
    folder_path = path_name
    with pd.ExcelWriter(output_excel_file, engine="xlsxwriter") as writer:
        for filename in os.listdir(folder_path):
            if filename.endswith('.csv'):
                file_path = os.path.join(folder_path, filename)
                ##################### CREATE ELEVATION TABLE FOR FULL CROSS SECTION GRAPH #####################
                df = pd.read_csv(file_path)
                df.to_excel(writer, sheet_name = filename, startrow=0, startcol=0, index=False)

                workbook = writer.book
                worksheet = writer.sheets[filename]
                
                
                global column_dict

                for index, column in enumerate(df.columns):
                    column_dict[f"{column}"] = index

                column_dict["First Available Column"] = df.shape[1]

                #CREATE Elev Ft COLUMN
                create_ft_converted_column(
                    worksheet = worksheet,
                    column_title="Elev Ft",
                    source_index=3,
                    target_index=column_dict["First Available Column"],
                    dataframe=df
                )
                 #CREATE Dist Ft COLUMN
                create_ft_converted_column(
                    worksheet = worksheet,
                    column_title="Dist Ft",
                    source_index=0,
                    target_index=column_dict["First Available Column"],
                    dataframe=df
                )
                create_depth_adjusted_column(
                    worksheet = worksheet,
                    column_title="Depth Ft",
                    source_index=column_dict["First Available Column"]-2,
                    target_index=column_dict["First Available Column"],
                    dataframe=df
                )

                #CREATE Historic Bankful Q CELL
                bankful_elevation_target_row = 23
                create_bankful_elevation_value_cells(
                    worksheet,
                    "Elevation of Historic Bankful Q Adjustable",
                    col_letter_to_index("AL"),
                    bankful_elevation_target_row
                )

                #CREATE True Bankful CELL
                create_bankful_elevation_value_cells(
                    worksheet,
                    "True Bankful Adjustable",
                    col_letter_to_index("AM"),
                    bankful_elevation_target_row
                )

                #CREATE Historic Bankful CELL
                create_bankful_elevation_value_cells(
                    worksheet,
                    "Historic Bankful Adjustable",
                    col_letter_to_index("AN"),
                    bankful_elevation_target_row
                )

                #CREATE Historic Bankful COLUMN
                create_bankful_elevation_columns(
                    worksheet = worksheet,
                    column_title="Historic Bankful",
                    source_index="$AN$25",
                    target_index=column_dict["First Available Column"],
                    dataframe=df
                )

                #CREATE True Bankful COLUMN
                create_bankful_elevation_columns(
                    worksheet = worksheet,
                    column_title="True Bankful",
                    source_index="$AM$25",
                    target_index=column_dict["First Available Column"],
                    dataframe=df
                )

                #ADD PADDING BETWEEN TABLES
                column_dict["First Available Column"] = column_dict["First Available Column"] + 1 
                ##################### CREATE TABLE FOR ZOOMED CROSS SECTION GRAPH #####################
                
                #CREATE ROW OFFSET COLUMN
                zoomed_rows = 45
                create_row_offset_column(
                    worksheet=worksheet,
                    column_title="Row Offset Column",
                    target_index=column_dict["First Available Column"],
                    dataframe=df,
                    number_zoomed_rows=zoomed_rows
                )

                #CREATE Dist M, X, Y, Elev M, Elev Ft, Dist Ft, and Depth Ft COLUMNS
                full_cross_section_column_list = ["Dist M", "X", "Y", "Elev M", "Elev Ft", "Dist Ft", "Depth Ft"]
                for column in full_cross_section_column_list:
                    create_zoomed_column(
                    worksheet=worksheet,
                    column_title=f"Zoomed {column}",
                    source_column=column_dict[column],
                    depth_column=column_dict["Depth Ft"],
                    row_offset_column=column_dict["Row Offset Column"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )

                #CREATE CELL WIDTH COLUMN
                create_zoomed_cell_width_column(
                worksheet=worksheet,
                column_title=f"Zoomed Cell Width",
                source_column=column_dict["Zoomed Dist Ft"],
                target_index=column_dict["First Available Column"],
                number_zoomed_rows=zoomed_rows
                )   
                #CREATE Av Cell Depth True Bankful
                create_zoomed_avg_cell_depth_true_bankful_column(
                worksheet=worksheet,
                column_title=f"Zoomed Av Cell Depth True Bankful",
                source_column=column_dict["Zoomed Depth Ft"],
                bankful_elevation_value_col=column_dict["True Bankful Adjustable"],
                bankful_elevation_value_row=bankful_elevation_target_row+2,
                target_index=column_dict["First Available Column"],
                number_zoomed_rows=zoomed_rows
                )         
                #CREATE Av Cell Depth Elevation Historic BF Q
                create_zoomed_avg_cell_depth_true_bankful_column(
                worksheet=worksheet,
                column_title=f"Zoomed Av Cell Depth Historic Bankful Q",
                source_column=column_dict["Zoomed Depth Ft"],
                bankful_elevation_value_col=column_dict["Elevation of Historic Bankful Q Adjustable"],
                bankful_elevation_value_row=bankful_elevation_target_row+2,
                target_index=column_dict["First Available Column"],
                number_zoomed_rows=zoomed_rows
                )     
                
                create_zoomed_bankful_column(
                    worksheet=worksheet,
                    column_title=f"Zoomed True Bankful Elevation",
                    source_column=column_dict["True Bankful Adjustable"],
                    source_row=bankful_elevation_target_row+2,
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_zoomed_Xca_column(
                    worksheet=worksheet,
                    column_title=f"Xca",
                    width_column=column_dict["Zoomed Cell Width"],
                    depth_column=column_dict["Zoomed Av Cell Depth True Bankful"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_zoomed_bankful_column(
                    worksheet=worksheet,
                    column_title=f"Zoomed Elevation of Historic Bankful Q Adjustable",
                    source_column=column_dict["Elevation of Historic Bankful Q Adjustable"],
                    source_row=bankful_elevation_target_row+2,
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_zoomed_2x_bankful_depth_column(
                    worksheet=worksheet,
                    column_title=f"Zoomed 2x True Bankful Depth",
                    source_column=column_dict["True Bankful Adjustable"],
                    source_row=bankful_elevation_target_row+2,
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )

                create_zoomed_Xca_stream_stats_column(
                    worksheet=worksheet,
                    column_title=f"Xca STREAMSTATS",
                    width_column=column_dict["Zoomed Cell Width"],
                    hist_bfq_column=column_dict["Zoomed Elevation of Historic Bankful Q Adjustable"],
                    xca_column = column_dict["Xca"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )

                create_zoomed_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Distance Cells Under Bankful Depth",
                    depth_column=column_dict["Zoomed Depth Ft"],
                    true_bankful_elevation_column=column_dict["Zoomed True Bankful Elevation"],
                    dist_column = column_dict["Zoomed Dist Ft"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_cleaned_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Cleaned Distance Cells Under Bankful Depth",
                    dist_cells_under_bf_col=column_dict["Distance Cells Under Bankful Depth"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )
                
                create_zoomed_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Distance Cells Under 2x Bankful Depth",
                    depth_column=column_dict["Zoomed Depth Ft"],
                    true_bankful_elevation_column=column_dict["Zoomed 2x True Bankful Depth"],
                    dist_column = column_dict["Zoomed Dist Ft"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_cleaned_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Cleaned Distance Cells Under 2x Bankful Depth",
                    dist_cells_under_bf_col=column_dict["Distance Cells Under 2x Bankful Depth"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )

                create_zoomed_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Distance Cells Under Historic Bankful Q",
                    depth_column=column_dict["Zoomed Depth Ft"],
                    true_bankful_elevation_column=column_dict["Zoomed Elevation of Historic Bankful Q Adjustable"],
                    dist_column = column_dict["Zoomed Dist Ft"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                    )
                
                create_cleaned_distance_cells_under_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Cleaned Distance Cells Under Historic Bankful Q",
                    dist_cells_under_bf_col=column_dict["Distance Cells Under Historic Bankful Q"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )

                create_trapezoid_from_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Trapezoide From Historice Bankful Q",
                    cleaned_dist_column=column_dict["Cleaned Distance Cells Under Historic Bankful Q"],
                    width_column = column_dict["Zoomed Cell Width"],
                    depth_column = column_dict["Zoomed Av Cell Depth Historic Bankful Q"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )

                create_trapezoid_from_bankful_depth(
                    worksheet=worksheet,
                    column_title=f"Trapezoide From True Bankful",
                    cleaned_dist_column=column_dict["Cleaned Distance Cells Under Bankful Depth"],
                    width_column = column_dict["Zoomed Cell Width"],
                    depth_column = column_dict["Zoomed Av Cell Depth Historic Bankful Q"],
                    target_index=column_dict["First Available Column"],
                    number_zoomed_rows=zoomed_rows
                )

                format_first_row(workbook, worksheet)
                
                generate_chart(
                    df=df,
                    workbook=workbook,
                    worksheet=worksheet,
                    chart_title="Cross Section From Left Bank",
                    x_axis_data=column_dict["Dist Ft"],
                    y_data_list=["Depth Ft", "Historic Bankful", "True Bankful"],
                    chart_row_index= 50,
                    num_chart_rows=len(df)
                )

                generate_chart(
                    df=df,
                    workbook=workbook,
                    worksheet=worksheet,
                    chart_title="Zoomed Cross Section From Left Bank",
                    x_axis_data=column_dict["Zoomed Dist Ft"],
                    y_data_list=["Zoomed Depth Ft", "Zoomed Elevation of Historic Bankful Q Adjustable", "Zoomed True Bankful Elevation"],
                    chart_row_index= 70,
                    num_chart_rows=zoomed_rows,
                    zoomed=True
                )
                
                
                print(f"generated {filename}")
        # for key in column_dict.keys():
        #     print(f"{key}")
######################## CELL POPULATING FUNCTIONS ################################

def num_to_excel_col(num):
    col_letter=""
    while num>=0:
        col_letter = chr(num % 26 + ord('A')) + col_letter
        num = num//26 - 1
    return col_letter

def col_letter_to_index(col_letter):
    """Convert Excel column letters to a zero-based numerical index."""
    col_index = 0
    for char in col_letter:
        col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
    return col_index - 1

def create_ft_converted_column(worksheet, column_title, source_index, target_index, dataframe):
    source_letter = num_to_excel_col(source_index)
    worksheet.write(0, target_index, column_title)
    
    global column_dict
    

    for row in range(len(dataframe)):
        formula = f"{source_letter}{row+2}*3.2808"
        worksheet.write_formula(row+1,target_index, formula)
    
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_bankful_elevation_value_cells(worksheet, column_title,  target_col, target_row):
    col_letter = num_to_excel_col(target_col)  # Convert column letters to numerical index
    worksheet.write(target_row, target_col,  column_title)  # Write the title
    formula = f"1.5"  # Formula reference using letters
    column_dict[column_title] = target_col
    worksheet.write_formula(target_row + 1,target_col,  formula)  # Write formula

def create_bankful_elevation_columns(worksheet, column_title, source_index, target_index, dataframe):
    worksheet.write(0, target_index, column_title)
    for row in range(len(dataframe)):
        formula = f"{source_index}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1
                
def create_depth_adjusted_column(worksheet, column_title, source_index, target_index, dataframe):
    source_letter = num_to_excel_col(source_index)
    worksheet.write(0, target_index, column_title)

    global column_dict

    for row in range(len(dataframe)):
        formula = f"{source_letter}{row+2}-MIN(${source_letter}$2:${source_letter}${len(dataframe)})"
        worksheet.write_formula(row+1,target_index, formula)

    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_row_offset_column(worksheet, column_title, target_index, dataframe, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    start_num = 0 - number_zoomed_rows//2
    col_letter = num_to_excel_col(target_index)
    for row in range(number_zoomed_rows - 1):
        if row == 0:
            worksheet.write(row + 1, target_index, start_num)
        else:
            formula = f"{col_letter}{row+1}+1"
            worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_column(worksheet, column_title, source_column, depth_column, row_offset_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    source_col_letter = num_to_excel_col(source_column)
    depth_column_letter = num_to_excel_col(depth_column)
    row_offset_column_letter = num_to_excel_col(row_offset_column)
    for row in range(number_zoomed_rows - 1):
        # Formula:   INDEX(B:B, MATCH(MIN($H:$H),$H:$H, 0) + $L3)
        formula = f"INDEX({source_col_letter}:{source_col_letter}, MATCH(MIN(${depth_column_letter}:${depth_column_letter}), ${depth_column_letter}:${depth_column_letter}, 0) + ${row_offset_column_letter}{row + 2})"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_cell_width_column(worksheet, column_title, source_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    source_col_letter = num_to_excel_col(source_column)
    for row in range(number_zoomed_rows - 2):
        # Formula:   =R4-R3
        formula = f"{source_col_letter}{row+3} - {source_col_letter}{row+2}"
        worksheet.write_formula(row+2,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_avg_cell_depth_true_bankful_column(worksheet, column_title, source_column, bankful_elevation_value_col, bankful_elevation_value_row,target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    source_col_letter = num_to_excel_col(source_column)
    bankful_elevation_value_col_letter = num_to_excel_col(bankful_elevation_value_col)
    for row in range(number_zoomed_rows - 2):
        # Formula:   =$AM$24-(S3+S4)/2
        formula = f"${bankful_elevation_value_col_letter}${bankful_elevation_value_row} - ({source_col_letter}{row+3} + {source_col_letter}{row+2})/2"
        worksheet.write_formula(row+2,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_bankful_column(worksheet, column_title, source_column, source_row, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    source_col_letter = num_to_excel_col(source_column)
    for row in range(number_zoomed_rows - 1):
        formula = f"${source_col_letter}${source_row}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_Xca_column(worksheet, column_title, width_column, depth_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    width_col_letter = num_to_excel_col(width_column)
    depth_col_letter = num_to_excel_col(depth_column)
    for row in range(number_zoomed_rows - 1):
        # Formula:  =T3*U3 
        formula = f"{width_col_letter}{row+2}*{depth_col_letter}{row+2}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_2x_bankful_depth_column(worksheet, column_title, source_column, source_row, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    source_col_letter = num_to_excel_col(source_column)
    for row in range(number_zoomed_rows - 1):
        formula = f"2*${source_col_letter}${source_row}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_Xca_stream_stats_column(worksheet, column_title, width_column, hist_bfq_column, xca_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    width_col_letter = num_to_excel_col(width_column)
    hist_bfq_col_letter = num_to_excel_col(hist_bfq_column)
    xca_col_letter = num_to_excel_col(xca_column)
    for row in range(number_zoomed_rows - 1):
        # Formula:  =T3*U3 
        formula = f"({width_col_letter}{row+2}*{hist_bfq_col_letter}{row+2}) - {xca_col_letter}{row+2}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_zoomed_distance_cells_under_bankful_depth(worksheet, column_title, depth_column, true_bankful_elevation_column, dist_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    depth_col_letter = num_to_excel_col(depth_column)
    true_bankful_elevation_col_letter = num_to_excel_col(true_bankful_elevation_column)
    dist_col_letter = num_to_excel_col(dist_column)
    for row in range(number_zoomed_rows - 1):
        # Formula:  =T3*U3 
        formula = f"IF({depth_col_letter}{row+2} >= {true_bankful_elevation_col_letter}{row+2}, NA(),1)*{dist_col_letter}{row+2}"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_cleaned_distance_cells_under_bankful_depth(worksheet, column_title, dist_cells_under_bf_col, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    start_num = 0 - number_zoomed_rows//2
    target_col_letter = num_to_excel_col(target_index)
    dist_cells_under_bf_col_letter = num_to_excel_col(dist_cells_under_bf_col)
    for index, row in enumerate(range(number_zoomed_rows - 1)):
        if (index) + start_num == 0:
            formula = f"{dist_cells_under_bf_col_letter}{row+2}"
            worksheet.write_formula(row+1,target_index, formula)
        elif (index) + start_num < 0:
            #Formula:   =IF(ISNUMBER(AC24),[@[Distance Cells Under Bankful Depth]],"#N/A")
            formula = f"IF(ISNUMBER({target_col_letter}{row+3}),{dist_cells_under_bf_col_letter}{row+2}, NA())"
            # formula = f"{index - start_num}"
            worksheet.write_formula(row+1,target_index, formula)
        elif(index) + start_num > 0:
             #Formula:   =IF(ISNUMBER(AC24),[@[Distance Cells Under Bankful Depth]],"#N/A")
            formula = f"IF(ISNUMBER({target_col_letter}{row+1}),{dist_cells_under_bf_col_letter}{row+2}, NA())"
            # formula = f"{index - start_num}"
            worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def create_trapezoid_from_bankful_depth(worksheet, column_title, cleaned_dist_column, width_column, depth_column, target_index, number_zoomed_rows):
    worksheet.write(0, target_index, column_title)
    target_col_letter = num_to_excel_col(target_index)
    cleaned_dist_col_letter = num_to_excel_col(cleaned_dist_column)
    width_column_letter = num_to_excel_col(width_column)
    depth_column_letter = num_to_excel_col(depth_column)
    for  row in range(number_zoomed_rows - 1):
        #Formula:  =IF(ISNUMBER([@[Cleaned Distnace Cells Elevation Historic BF ]]),V4*(T4),0)
        formula = f"IF(ISNUMBER({cleaned_dist_col_letter}{row+2}),{width_column_letter}{row+2}*{depth_column_letter}{row+2},0)"
        worksheet.write_formula(row+1,target_index, formula)
    column_dict[column_title] = target_index
    column_dict["First Available Column"] = target_index + 1

def format_first_row(workbook, worksheet):
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })
    worksheet.set_row(0, None, header_format)

def generate_chart(df, workbook,  worksheet, chart_title, x_axis_data,  y_data_list, chart_row_index, num_chart_rows, zoomed = False):
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
    start_row = 1
    end_row = num_chart_rows

    x_axis_col_letter = num_to_excel_col(x_axis_data)
    x_axis_index = col_letter_to_index(x_axis_col_letter)

    if zoomed:
        min_index = df['Elev M'].idxmin()
        # print("DF Length: ", len(df), "MIN INDEX: ", min_index, "Proposed Distance: ",df.loc[min_index , 'Dist M'])
        min_x_index = max(0, min_index - num_chart_rows//2)
        max_x_index = min(len(df) - 1, min_index + num_chart_rows//2)

        

        min_x_value = df.loc[min_x_index , 'Dist M'] * 3.2808
        max_x_value = df.loc[max_x_index, 'Dist M'] * 3.2808
        # print("MIN X Value: ",min_x_value,"MAX X VALUE: ",max_x_value )

    for series in y_data_list:
        y_series_col_letter = num_to_excel_col(column_dict[series])
        y_series_index = col_letter_to_index(y_series_col_letter)
        chart.add_series({
            'name':       [worksheet.name, 0, y_series_index],  # Header for the legend
            'categories': [worksheet.name, start_row, x_axis_index, end_row, x_axis_index],  # X values
            'values':     [worksheet.name, start_row, y_series_index, end_row, y_series_index],  # Y values
            # 'marker':     {'type': 'circle'}  # Optional: Custom marker type
        })
    

    chart.set_title({'name': f"{chart_title}"})
    if zoomed:
        chart.set_x_axis({
        'name': 'Distance Ft',
        'min': min_x_value,       # Minimum x-axis value
        'max': max_x_value,      # Maximum x-axis value
        })
    else:
        chart.set_x_axis({'name': 'Distance Ft'})
    chart.set_y_axis({'name': 'Depth Ft'})
    chart.set_legend({'position': 'bottom'})

    # Insert the chart into the worksheet
    worksheet.insert_chart(f'L{chart_row_index}', chart,{
        'x_scale': 3,  # 50% wider than default
        'y_scale': 1   # 20% taller than default
    })

########################################## ---- MAIN ----- ###########################################

if __name__ ==  "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel sheets to CSV files")
    parser.add_argument("file_path", help="Path to the excel file")
    parser.add_argument(
        "--output_excel_file",
        default = "converted_excel_file.xlsx",
        help="Output folder for csv files (defaul: 'converted_excel_file')"
    )

    args = parser.parse_args()
    sheet_generator(args.file_path, args.output_excel_file)