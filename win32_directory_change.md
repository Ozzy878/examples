#### Change to CWD using Pathlib
    output_file_name = "the name of your file"
    output_folder = "S://Path/to/your/desired/directory"
    output_file = Path('{}/{}.xls'.format(output_folder, output_file_name))

##### Save the file with the following string
    wb.SaveAs(str(output_file), FileFormat=56, ConflictResolution=2)