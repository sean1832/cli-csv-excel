import pandas as pd
import sys
import os

try:
    # check if arguments are given correctly
    if len(sys.argv) < 2:
        print("Wrong number of arguments!")
        print("Usage: csv-excel <input_file_path> [output_file_name]")
        print(r"Example: csv-excel C:\path\to\your\inputfile.csv [outputfilename]")
        exit(1)
    else:
        # convert csv to xlsx
        print("Converting CSV to XLSX...")

        # parse directory from input file path
        directory = os.path.dirname(sys.argv[1])
        input_filename_without_extension = os.path.splitext(os.path.basename(sys.argv[1]))[0]

        # Determine the output file path
        if len(sys.argv) == 3:
            if os.sep in sys.argv[2]:
                # if output file path is given, use it
                output_filepath = sys.argv[2] + '.xlsx'
            else:
                # if only output file name is given, use the input file path
                output_filepath = os.path.join(directory, sys.argv[2] + '.xlsx')
        else:
            # if no output file name is given, use the input file name
            output_filepath = os.path.join(directory, input_filename_without_extension + '.xlsx')

        # Read CSV
        data_xls = pd.read_csv(sys.argv[1], index_col=0)
        # Write to XLSX
        data_xls.to_excel(output_filepath, index=True)

        print(f"File saved to {output_filepath}")
        print("Done!")

except Exception as e:
    print("Failed to convert CSV to XLSX: " + str(e))
    exit(1)
