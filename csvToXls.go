package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Define flags
	fileFlag := flag.String("f", "", "Path to a single CSV file to convert")
	dirFlag := flag.String("d", "", "Path to a directory containing CSV files to convert")
	singleFileFlag := flag.Bool("s", false, "In directory mode, create a single Excel file with multiple sheets instead of separate files")

	// Customize help message
	flag.Usage = customHelp

	// Parse flags
	flag.Parse()

	// If help was explicitly requested, show it and exit
	for _, arg := range os.Args[1:] {
		if arg == "-h" || arg == "--help" {
			customHelp()
			os.Exit(0)
		}
	}

	// Verify that at least one of the mandatory flags is specified
	if *fileFlag == "" && *dirFlag == "" {
		fmt.Println("Error: You must specify either -f (file) or -d (directory)")
		customHelp()
		os.Exit(1)
	}

	// Verify that both flags are not specified together
	if *fileFlag != "" && *dirFlag != "" {
		fmt.Println("Error: Specify either -f or -d, not both")
		os.Exit(1)
	}

	// Process based on the specified flag
	if *fileFlag != "" {
		// Single file mode
		err := processFile(*fileFlag, "")
		if err != nil {
			fmt.Printf("Error during file conversion: %v\n", err)
			os.Exit(1)
		}
	} else {
		// Directory mode
		if *singleFileFlag {
			// Single file with multiple sheets mode
			err := processDirectoryToSingleFile(*dirFlag)
			if err != nil {
				fmt.Printf("Error during directory conversion: %v\n", err)
				os.Exit(1)
			}
		} else {
			// Separate files mode
			err := processDirectory(*dirFlag)
			if err != nil {
				fmt.Printf("Error during directory conversion: %v\n", err)
				os.Exit(1)
			}
		}
	}
}

// Custom function for help
func customHelp() {
	fmt.Println("Usage: csvtoxls [options]")
	fmt.Println("\nOptions:")
	fmt.Println("  -f file.csv     Converts a single CSV file to XLSX")
	fmt.Println("  -d directory    Converts all CSV files in the specified directory")
	fmt.Println("  -s              In directory mode, creates a single Excel file with multiple sheets")
	fmt.Println("                  instead of creating one XLSX file per CSV")
	fmt.Println("  -h, --help      Shows this help message")
	fmt.Println("\nExamples:")
	fmt.Println("  csvtoxls -f data.csv                   # Converts a single file")
	fmt.Println("  csvtoxls -d ./data                     # Converts all CSVs to separate files")
	fmt.Println("  csvtoxls -d ./data -s                  # Converts all CSVs to a single Excel file")
	fmt.Println("\nNotes:")
	fmt.Println("  - The default separator is semicolon (;)")
	fmt.Println("  - Quotes are removed from values")
	fmt.Println("  - Column widths are automatically adjusted to fit content")
	fmt.Println("  - Existing files will be overwritten without warning")
}

// Process a single CSV file
func processFile(csvFilePath, sheetName string) error {
	// Verify that the file exists
	if _, err := os.Stat(csvFilePath); os.IsNotExist(err) {
		return fmt.Errorf("file %s does not exist", csvFilePath)
	}

	// Verify that the file has a .csv extension
	if !strings.HasSuffix(strings.ToLower(csvFilePath), ".csv") {
		return fmt.Errorf("file %s is not a CSV file", csvFilePath)
	}

	// If no sheet name is specified, use the file name
	if sheetName == "" {
		// Extract the file name without extension
		baseName := filepath.Base(csvFilePath)
		sheetName = strings.TrimSuffix(baseName, filepath.Ext(baseName))

		// Make sure the sheet name is valid for Excel (max 31 characters, no special characters)
		if len(sheetName) > 31 {
			sheetName = sheetName[:31]
		}
		// Replace invalid characters with underscores
		sheetName = sanitizeSheetName(sheetName)
	}

	// Create name for the Excel file
	xlsxFilePath := strings.TrimSuffix(csvFilePath, filepath.Ext(csvFilePath)) + ".xlsx"

	// Create a new Excel file
	f := excelize.NewFile()

	// Get the default sheet name
	defaultSheet := f.GetSheetName(0) // Usually "Sheet1"

	// Create a new sheet with the appropriate name
	f.NewSheet(sheetName)

	// Convert the CSV content
	columnWidths, err := convertCSVtoSheet(csvFilePath, f, sheetName)
	if err != nil {
		return fmt.Errorf("conversion failed for %s: %v", csvFilePath, err)
	}

	// Adjust column widths to fit content
	adjustColumnWidths(f, sheetName, columnWidths)

	// Set the active sheet
	index, _ := f.GetSheetIndex(sheetName)
	f.SetActiveSheet(index)

	// Delete the default sheet after setting the active sheet
	f.DeleteSheet(defaultSheet)

	// Save the Excel file
	err = f.SaveAs(xlsxFilePath)
	if err != nil {
		return fmt.Errorf("error saving Excel file %s: %v", xlsxFilePath, err)
	}

	fmt.Printf("Conversion completed: %s -> %s\n", csvFilePath, xlsxFilePath)
	return nil
}

// Process all CSV files in a directory (separate files)
func processDirectory(dirPath string) error {
	// Verify that the directory exists
	if _, err := os.Stat(dirPath); os.IsNotExist(err) {
		return fmt.Errorf("directory %s does not exist", dirPath)
	}

	// Counters for statistics
	var successCount, failCount int

	// Visit all files in the directory
	err := filepath.WalkDir(dirPath, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}

		// Skip directories
		if d.IsDir() {
			return nil
		}

		// Process only CSV files
		if strings.HasSuffix(strings.ToLower(path), ".csv") {
			err := processFile(path, "")
			if err != nil {
				fmt.Printf("ERROR: %v\n", err)
				failCount++
			} else {
				successCount++
			}
		}

		return nil
	})

	if err != nil {
		return fmt.Errorf("error scanning directory: %v", err)
	}

	// Print statistics
	fmt.Printf("\nSummary: %d files successfully converted, %d failed\n", successCount, failCount)

	if successCount == 0 && failCount == 0 {
		fmt.Println("No CSV files found in the directory")
	}

	return nil
}

// Process all CSV files in a directory (single file with multiple sheets)
func processDirectoryToSingleFile(dirPath string) error {
	// Verify that the directory exists
	if _, err := os.Stat(dirPath); os.IsNotExist(err) {
		return fmt.Errorf("directory %s does not exist", dirPath)
	}

	// Name of the output Excel file
	dirName := filepath.Base(dirPath)
	xlsxFilePath := filepath.Join(dirPath, dirName+".xlsx")

	// Create a new Excel file
	f := excelize.NewFile()

	// Get the default sheet name
	defaultSheet := f.GetSheetName(0) // Usually "Sheet1"

	// Counters for statistics
	var successCount, failCount int
	var firstSheet string

	// Collect all CSV files
	var csvFiles []string
	err := filepath.WalkDir(dirPath, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}

		// Skip directories
		if d.IsDir() {
			return nil
		}

		// Collect only CSV files
		if strings.HasSuffix(strings.ToLower(path), ".csv") {
			csvFiles = append(csvFiles, path)
		}

		return nil
	})

	if err != nil {
		return fmt.Errorf("error scanning directory: %v", err)
	}

	// Check if there are CSV files
	if len(csvFiles) == 0 {
		fmt.Println("No CSV files found in the directory")
		return nil
	}

	// Map to keep track of sheet names (to avoid duplicates)
	sheetNames := make(map[string]bool)

	// Process all CSV files
	for _, csvFilePath := range csvFiles {
		// Extract the file name without extension to use as sheet name
		baseName := filepath.Base(csvFilePath)
		sheetName := strings.TrimSuffix(baseName, filepath.Ext(baseName))

		// Make sure the sheet name is valid for Excel (max 31 characters)
		if len(sheetName) > 31 {
			sheetName = sheetName[:31]
		}

		// Sanitize the sheet name
		sheetName = sanitizeSheetName(sheetName)

		// Handle duplicate names
		originalName := sheetName
		counter := 1
		for sheetNames[sheetName] {
			// If the name already exists, add a number
			suffix := fmt.Sprintf("_%d", counter)

			// Make sure the name with the suffix doesn't exceed 31 characters
			if len(originalName)+len(suffix) > 31 {
				sheetName = originalName[:31-len(suffix)] + suffix
			} else {
				sheetName = originalName + suffix
			}

			counter++
		}

		// Register the sheet name
		sheetNames[sheetName] = true

		// Create a new sheet
		_, err := f.NewSheet(sheetName)
		if err != nil {
			fmt.Printf("ERROR: Unable to create sheet %s: %v\n", sheetName, err)
			failCount++
			continue
		}

		// Save the name of the first sheet to set it as active
		if firstSheet == "" {
			firstSheet = sheetName
		}

		// Convert the CSV content
		columnWidths, err := convertCSVtoSheet(csvFilePath, f, sheetName)
		if err != nil {
			fmt.Printf("ERROR: %v\n", err)
			failCount++
		} else {
			// Adjust column widths to fit content
			adjustColumnWidths(f, sheetName, columnWidths)
			fmt.Printf("Sheet '%s' created from %s\n", sheetName, csvFilePath)
			successCount++
		}
	}

	// Set the first sheet as active (if it exists)
	if firstSheet != "" {
		index, _ := f.GetSheetIndex(firstSheet)
		f.SetActiveSheet(index)

		// Delete the default sheet after setting the active sheet
		f.DeleteSheet(defaultSheet)
	}

	// Save the Excel file
	err = f.SaveAs(xlsxFilePath)
	if err != nil {
		return fmt.Errorf("error saving Excel file %s: %v", xlsxFilePath, err)
	}

	// Print statistics
	fmt.Printf("\nExcel file created: %s\n", xlsxFilePath)
	fmt.Printf("Summary: %d sheets successfully created, %d failed\n", successCount, failCount)

	return nil
}

// Convert a CSV to an Excel sheet and return column widths
func convertCSVtoSheet(csvFilePath string, f *excelize.File, sheetName string) (map[int]int, error) {
	// Open the CSV file
	csvFile, err := os.Open(csvFilePath)
	if err != nil {
		return nil, fmt.Errorf("unable to open CSV file: %v", err)
	}
	defer csvFile.Close()

	// Create a new CSV reader with appropriate settings
	reader := csv.NewReader(csvFile)
	reader.Comma = ';'             // Set the separator as semicolon
	reader.FieldsPerRecord = -1    // Allow variable number of fields per row
	reader.LazyQuotes = true       // Handle quotes more flexibly
	reader.TrimLeadingSpace = true // Remove leading spaces

	// Map to track the maximum width of each column
	columnWidths := make(map[int]int)

	// Read and process the CSV row by row
	rowIndex := 1
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("error reading CSV at row %d: %v", rowIndex, err)
		}

		// Insert data into the Excel sheet
		for colIndex, value := range record {
			// Remove quotes at the beginning and end
			value = strings.TrimPrefix(value, "\"")
			value = strings.TrimSuffix(value, "\"")

			// Convert indices to cell name (A1, B1, etc.)
			cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex)
			if err != nil {
				return nil, fmt.Errorf("error converting coordinates: %v", err)
			}

			// Set the value in the cell
			if err := f.SetCellValue(sheetName, cellName, value); err != nil {
				return nil, fmt.Errorf("error setting cell value: %v", err)
			}

			// Update the maximum width for this column
			// Add a bit of padding (1.2 multiplier) for better appearance
			valueWidth := int(float64(utf8.RuneCountInString(value)) * 1.2)
			if valueWidth > columnWidths[colIndex] {
				columnWidths[colIndex] = valueWidth
			}
		}
		rowIndex++
	}

	return columnWidths, nil
}

// Adjust column widths to fit content
func adjustColumnWidths(f *excelize.File, sheetName string, columnWidths map[int]int) {
	// Set minimum and maximum width limits
	const (
		minWidth = 8
		maxWidth = 100
	)

	// Adjust each column width
	for colIndex, width := range columnWidths {
		// Apply minimum and maximum constraints
		if width < minWidth {
			width = minWidth
		} else if width > maxWidth {
			width = maxWidth
		}

		// Convert column index to column name (A, B, C, etc.)
		colName, _ := excelize.ColumnNumberToName(colIndex + 1)

		// Set the column width
		f.SetColWidth(sheetName, colName, colName, float64(width))
	}
}

// Sanitize the sheet name by removing invalid characters
func sanitizeSheetName(name string) string {
	// Characters not allowed in Excel sheet names: [ ] * ? / \ : '
	invalidChars := []string{"[", "]", "*", "?", "/", "\\", ":", "'"}
	result := name

	for _, char := range invalidChars {
		result = strings.ReplaceAll(result, char, "_")
	}

	// Make sure the name is not empty
	if result == "" {
		result = "Sheet"
	}

	return result
}
