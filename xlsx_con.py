import os
import glob
import pandas as pd

def print_banner():
    banner = """
   ██╗  ██╗██╗     ███████╗██╗  ██╗     ██████╗ ██████╗ ███╗   ██╗██╗   ██╗███████╗██████╗ ███████╗██╗ ██████╗ ███╗   ██╗    
   ╚██╗██╔╝██║     ██╔════╝╚██╗██╔╝    ██╔════╝██╔═══██╗████╗  ██║██║   ██║██╔════╝██╔══██╗██╔════╝██║██╔═══██╗████╗  ██║    
    ╚███╔╝ ██║     ███████╗ ╚███╔╝     ██║     ██║   ██║██╔██╗ ██║██║   ██║█████╗  ██████╔╝███████╗██║██║   ██║██╔██╗ ██║    
    ██╔██╗ ██║     ╚════██║ ██╔██╗     ██║     ██║   ██║██║╚██╗██║╚██╗ ██╔╝██╔══╝  ██╔══██╗╚════██║██║██║   ██║██║╚██╗██║    
██╗██╔╝ ██╗███████╗███████║██╔╝ ██╗    ╚██████╗╚██████╔╝██║ ╚████║ ╚████╔╝ ███████╗██║  ██║███████║██║╚██████╔╝██║ ╚████║    
╚═╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝     ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝  ╚═══╝  ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝ ╚═════╝ ╚═╝  ╚═══╝  

    """
    print(banner)

def convert_xlsx_to_csv(directory):
    xlsx_files = glob.glob(os.path.join(directory, '*.xlsx'))
    
    csv_output_directory = os.path.join(directory, 'csv_output')
    os.makedirs(csv_output_directory, exist_ok=True)
    
    for file in xlsx_files:
        df = pd.read_excel(file)
        
        csv_file_name = os.path.splitext(os.path.basename(file))[0] + '.csv'
        csv_file_path = os.path.join(csv_output_directory, csv_file_name)
        
        df.to_csv(csv_file_path, index=False)
        print(f"Converted {file} to {csv_file_path}")
        
        convert_csv_to_text(csv_file_path)

    print(f"Conversion complete. {len(xlsx_files)} .xlsx files converted to .csv.")

def convert_csv_to_text(csv_file_path):
    df = pd.read_csv(csv_file_path)
    
    text_file_name = os.path.splitext(os.path.basename(csv_file_path))[0] + '.txt'
    text_file_path = os.path.splitext(csv_file_path)[0] + '_table.txt'
    
    with open(text_file_path, 'w') as f:
        columns = df.columns.tolist()
        
        col_widths = [max(len(str(item)) for item in df[col]) for col in columns]
        col_widths = [max(width, len(col)) for width, col in zip(col_widths, columns)]  
        
        format_str = ' | '.join(f'{{:<{width}}}' for width in col_widths)
        
        f.write(format_str.format(*columns) + '\n')
        f.write('-' * (sum(col_widths) + 3 * (len(columns) - 1)) + '\n')  
        
        for index, row in df.iterrows():
            f.write(format_str.format(*row) + '\n')
    
    print(f"Converted {csv_file_path} to {text_file_path}")

if __name__ == "__main__":
    print_banner()
   
    directory = input("Enter the directory to search for .xlsx files: ")
    
    convert_xlsx_to_csv(directory)