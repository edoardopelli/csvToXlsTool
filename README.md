# CSV to XLS Converter

The CSV to XLS Converter is a simple command-line tool written in Go that allows you to convert CSV files into Excel XLSX spreadsheets. It leverages the [excelize](https://github.com/xuri/excelize) library to generate Excel files.

## Features

- **CSV to Excel conversion:** Transform CSV data into an XLSX file for easy viewing and analysis.
- **Command-line operation:** Run the tool from the terminal with minimal setup.
- **Dependency management:** Automatically handles dependencies via Go modules.

## Prerequisites

- [Go](https://golang.org/dl/) must be installed on your system. You can check your installation with:
  ```bash
  go version

Installation
	1.	Clone the repository:

git clone <repository-url>
cd <repository-directory>

Replace <repository-url> with the URL of your repository and <repository-directory> with the name of the cloned folder.

	2.	Initialize the Go module:
If a go.mod file does not already exist, initialize one with:

go mod init csvtoxls


	3.	Download dependencies:
Use the following command to download and tidy up the dependencies:

go mod tidy



Usage

Building the Tool

To compile the program, run:

go build csvToXls.go

This command generates an executable file (csvToXls on Unix-like systems or csvToXls.exe on Windows).

Alternatively, you can compile and run the program directly with:

go run csvToXls.go

Running the Tool

Once compiled, run the executable from the terminal:
	•	On Unix/macOS:

./csvToXls


	•	On Windows:

csvToXls.exe



The tool expects a CSV file as input and converts it to an Excel XLSX format. Depending on your implementation, you might need to supply the input file path and an output destination via command-line arguments.

Example Command

For instance, if your tool supports command-line parameters:

./csvToXls -input=data.csv -output=data.xlsx

Adjust the parameters (-input and -output) to match your program’s implementation.

Troubleshooting
	•	Import Cycle Error:
If you encounter an error such as:

package command-line-arguments
    imports github.com/xuri/excelize/v2 from csvToXls.go: import cycle not allowed

This is typically due to a module name conflict in your go.mod file. Ensure the module name declared in go.mod does not conflict with any imported packages (e.g., change it to module csvtoxls).

	•	Missing Dependencies:
Running go mod tidy should resolve any missing package dependencies by automatically downloading them.

Contributing

Contributions are welcome! Feel free to fork the repository and submit pull requests with improvements and bug fixes.

License

Distributed under the BSD 2-Clause License. See LICENSE for more information.

Contact

For any questions, issues, or feature requests, please open an issue in the repository or contact the project maintainer.

