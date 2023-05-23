# 📝 Presentation Mapper

The Presentation Mapper is a Python tool that allows you to easily replace placeholders in a PowerPoint presentation with data from an Excel file. It uses a mapping file to define the placeholders and their corresponding data sources by cell references. If a string is provided in stead of a cell reference, the value will be used directly to replace the corresponding placeholder.

## 🚀 Add your source files

To use the Presentation Mapper, follow these steps:

Clone the repository to your local machine.
```
git clone https://github.com/your_username/presentation-mapper.git
```

Install the required dependencies by running:
```
pip install -r requirements.txt
```

Create or move PowerPoint presentation with placeholders in the format 
{{placeholder_name}} into the [source_files](./source_files/) subdirectory

Also create or move an excel spreadsheet you want to extract values from into the [source_files](./source_files/) subdirectory

>Note: The scripts only expect ONE presentation and ONE excel file to be in the [source_files](./source_files/) subdirectory.

## 🚀 Generate a mapping file

Generate a mapping file using the command

```
make mapping
```

This script takes the path to a PowerPoint presentation as input and generates a mapping file based on the placeholders found in the presentation.

It does this by inspecting your `.pptx` file in the [source_files](./source_files/) directory and finding all placeholders with the {{placeholder_format}}, adding them as keys in the [`./source_files`](./source_files/mapping.yml) file.

## 🚀 Fill in the mapping file

The mapping file will initially have `null` as values for all placeholder keys. Specify the values for each one from the `.xlsx` file that you would like to use in the presentation. These should be a reference to both a sheet and a cell e.g. "Sheet1!A3"

>Note: It is important that you follow the format "{SheetName}!{CellCoords}" exactly, or the script will not find the correct cell.

>Double Note: If you don't want to specify a cell, you can specify any string (without a ! character) istead, and this will be used in stead of a cell value at substitution time e.g. "My Presentation"

## 🚀 Apply the mapping

Apply your specified mapping by using the command:

```
make presentation
```

This takes the template `.pptx` file, replaces all {{placeholder}} instances within with their corresponding values from the spreadsheet, as specified in the [`source_files/mapping.yml`](./source_files/mapping.yml) file (created in the previous step).


## Rinse & repeat!

Transfer the result out of the directory. If you need to do it again, simply start these instructions again from the beginning.

> Note: Extra for experts. If you find yourself using the same template and mappings multiple times, it probably makes sense to save these in an appropriate subdirectory under [`.source_files`](./source_files/). The script only looks for the `./source_files/***.pptx` and `./source_files/***.xlsx` paths. So you can move your templates into this path to use them, then move them back to the template subdirectory for next time.

## 📂 Repository Structure

[`./scripts/execute_mapping.py`](./scripts/execute_mapping.py)
: the main script that executes the mapping process.
make_mapping.py
: a script that generates a mapping file based on the placeholders found in a PowerPoint presentation.
utils.py
: utility functions used by the other scripts.
example_mapping.yml
: an example mapping file, demonstrating the format. Note that the actual mapping file will be named `mapping.yml`
requirements.txt
: a list of required Python packages to run the scripts
README.md
: this file!
📝 Creating a Mapping File

To create a mapping file, you can use the 
make_mapping.py
 script. 

## Example:

Copy codepython make_mapping.py presentation.pptx mapping.yml
The resulting mapping file will be saved in the current directory with the name 
mapping.yml
.

## 📝 Defining Data Sources

The mapping file defines the placeholders in the presentation and their corresponding data sources. Each placeholder is defined as a key in the mapping file, and the value is the data source for that placeholder.

There are two types of data sources:

Direct Value: a string that will be used directly to replace the placeholder.
Cell Reference: a reference to a cell in an Excel file, in the format 
worksheet_name!cell_address
.
## 📝 Handling Errors

The Presentation Mapper includes error handling for common issues, such as missing files, invalid filenames, and empty cells in Excel. If an error occurs, the script will print an error message and exit.

🎉 Congratulations! You now know how to use the Presentation Mapper to easily replace placeholders in a PowerPoint presentation with data from an Excel file.