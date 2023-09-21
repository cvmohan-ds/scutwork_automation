import os
import shutil
import argparse


def move(in_dir, out_dir, doc_dir, xl_dir):
    """
    This utility moves files of given extension.
    Objective: Move files with given extension: .docx or .xlsx present in subdirectories of given directory.
    example:
    processed_info
        - info1
            - info1_formatted.docx
            - info1_data.xlsx
            - info1_log.log
        - info2
            - info2_formatted.docx
            - info2_data.xlsx
            - info2_log.log
        ...
        - info1003
            - info1003_formatted.docx
            - info1003_data.xlsx
            - info1003_log.log
    We need to move docx and xlsx files to two different folders
    processed_info_formatted - This is for word report - docx
    processed_info_data - This is the raw data for the report - xlsx
    """
    print("started the Process!!!")
    if os.path.isdir(in_dir):
        dirs = os.listdir(in_dir)
        check_and_create_out_dirs(out_dir, doc_dir, xl_dir)
        for directory in dirs:
            print(f"Moving {directory} content")
            for file_name in os.listdir(f"{in_dir}/{directory}"):
                if file_name.endswith(".docx"):
                    shutil.move(os.path.join(f"{in_dir}/{directory}", file_name), f"{out_dir}/{doc_dir}")
                    print("xlsx files moved")
                elif file_name.endswith(".xlsx"):
                    shutil.move(os.path.join(f"{in_dir}/{directory}", file_name), f"{out_dir}/{xl_dir}")
                    print("word File moved")
            print(f"Done with {directory}")



def check_and_create_out_dirs(out_folder, docs_folder, xls_folder):
    """_summary_
        This Creates the output directories if they are not available.
        similar to linunx command line mkdir -p <dir1/dir2> 
        creates both dir1 and dir2 if they do not exist.
    Args:
        out_folder (_type_): _description_
        docs_folder (_type_): _description_
        xls_folder (_type_): _description_
    """
    # check if output directory exists
    if os.path.isdir(out_folder):
        if os.path.isdir(f"{out_folder}/{docs_folder}"):
            print(" Word output directory exists")
        else:
            try:
                os.makedirs(f"{out_folder}/{docs_folder}", exist_ok=True)
            except OSError as os_error:
                print(" Word Output Directory cannot be created :", os_error)
        if os.path.isdir(f"{out_folder}/{xls_folder}"):
            print(" Excel output directory exists")
        else:
            try:
                os.makedirs(f"{out_folder}/{xls_folder}", exist_ok=True)
            except OSError as os_error:
                print(" Excel Output Directory cannot be created :", os_error)
    else:
        # if not create the directory
        try:
            os.makedirs(f"{out_folder}/{docs_folder}", exist_ok=True)
            os.makedirs(f"{out_folder}/{xls_folder}", exist_ok=True)
            print("Directories Created Successully: ")
        except OSError as os_error:
            print("Directory cannot be created :", os_error)



if __name__ == "__main__":
    print("Scutwork Automated via Python: Moving files to respective folders")
    parser = argparse.ArgumentParser("Scutwork Automation")
    parser.add_argument('-i', '--inputdir', help='Absolute path of input directory', required=True)
    parser.add_argument("-o", "--outputdir", help="Absolute path of output directory",
                        required=True)
    parser.add_argument("-d", "--subdirdocx", help=" Name of subdirectory for docx files",
                        required=False, default="word_info")
    parser.add_argument("-x", "--subdirxlsx", help="Name of subdirectory for xlsx files",
                        required=False, default="excel_info")
    options = parser.parse_args()
    input_dir = options.inputdir
    output_dir = options.outputdir
    word_dir = options.subdirdocx
    excel_dir = options.subdirxlsx
    move(input_dir, output_dir, word_dir, excel_dir)
