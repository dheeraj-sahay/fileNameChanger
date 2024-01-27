```
# CODAI
> main.go

```go
package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"strconv"
	"github.com/joho/godotenv"
	"github.com/360EntSecGroup-Skylar/excelize"
)

// changeFileName function takes old and new names and renames the file
func changeFileName(oldName string, newName string) error {
	err := os.Rename(oldName, newName)
	if err != nil {
		return err
	}
	return nil
}

// createExcelSheet function creates an Excel sheet with old and new names
func createExcelSheet(directory string, fileDetails [][]string) error {
	// Create a new Excel file
	xlsx := excelize.NewFile()

	// Set value for each cell in the Excel sheet
	for rowIndex, row := range fileDetails {
		for colIndex, value := range row {
			cell := excelize.ToAlphaString(colIndex + 1) + fmt.Sprint(rowIndex+1)
			xlsx.SetCellValue("Sheet1", cell, value)
		}
	}

	// Save the Excel file in the same directory
	err := xlsx.SaveAs(filepath.Join(directory, "file_rename_details.xlsx"))
	if err != nil {
		return err
	}
	return nil
}

// readSequenceFromEnv reads the current sequence from the environment variable
func readSequenceFromEnv() (int, error) {
	// Load environment variables from .env file
	err := godotenv.Load()
	if err != nil {
		return 0, err
	}

	// Get the sequence from the environment variable
	sequenceStr, exists := os.LookupEnv("SEQUENCE")
	if !exists {
		return 0, fmt.Errorf("SEQUENCE environment variable not found")
	}

	sequence, err := strconv.Atoi(sequenceStr)
	if err != nil {
		return 0, err
	}

	return sequence, nil
}

// writeSequenceToEnv writes the updated sequence to the environment variable
func writeSequenceToEnv(sequence int) error {
	// Update the environment variable
	err := os.Setenv("SEQUENCE", strconv.Itoa(sequence))
	if err != nil {
		return err
	}

	return nil
}

func main() {
	// Get the current working directory
	currentDir, err := os.Getwd()
	if err != nil {
		fmt.Println("Error getting current directory:", err)
		return
	}

	// Get all file names in the directory
	fileNames, err := filepath.Glob(filepath.Join(currentDir, "*"))
	if err != nil {
		fmt.Println("Error getting file names:", err)
		return
	}

	// Initialize a slice to store file details for the Excel sheet
	var fileDetails [][]string

	// Read the current sequence from the environment variable
	sequence, err := readSequenceFromEnv()
	if err != nil {
		fmt.Println("Error reading sequence from environment variable:", err)
		return
	}

	// Loop through each file
	for _, fileName := range fileNames {
		// Generate new file name with an incrementing sequence
		sequence++
		newFileName := fmt.Sprintf("file_%d%s", sequence, filepath.Ext(fileName))

		// Rename the file
		err := changeFileName(fileName, newFileName)
		if err != nil {
			fmt.Println("Error renaming file:", err)
			return
		}

		// Append file details to the slice
		fileDetails = append(fileDetails, []string{fileName, newFileName})
	}

	// Write the updated sequence back to the environment variable
	err = writeSequenceToEnv(sequence)
	if err != nil {
		fmt.Println("Error writing sequence to environment variable:", err)
		return
	}

	// Create an Excel sheet with file details
	err = createExcelSheet(currentDir, fileDetails)
	if err != nil {
		fmt.Println("Error creating Excel sheet:", err)
		return
	}

	fmt.Println("File names changed successfully. Excel sheet created.")
}
```

### File Tree

```
/
├── main.go
└── .env
```  
DONE.