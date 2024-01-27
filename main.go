package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// change File Name function
func changeFileName(oldName string, newName string) error {
	err := os.Rename(oldName, newName)
	if err != nil {
		return err
	}
	return nil
}

// create Excel Sheet with old file name and new file name
func createExecelSheet(directory string, fileDetails [][]string) error {
	// create a new Excel File
	xlsx := excelize.NewFile()
	// Set value of each cell in excel sheet
	for rowIndex, row := range fileDetails {
		for colIndex, value := range row {
			cell := excelize.ToAlphaString(colIndex) + fmt.Sprint(rowIndex+1)
			xlsx.SetCellValue("Sheet1", cell, value)
		}
	}

	// save the excel file in the same directory
	err := xlsx.SaveAs(filepath.Join(directory, "file_rename_details.xlsx"))
	if err != nil {
		return err
	}

	return nil
}

// read from constant file
func readVariablesFromConstent() (string, int, error) {
	fmt.Println("Entered in read function!")
	f, err := os.Open("Constant.txt")
	if err != nil {
		fmt.Println("Error opening the constent file:", err)
		return "", 0, err
	}

	defer f.Close()

	scanner := bufio.NewScanner(f)
	var pattern string = ""
	var sequence int = 0
	for scanner.Scan() {
		line := scanner.Text()
		comps := strings.Split(line, "=")
		if "PATTERN" == comps[0] {
			pattern = comps[1]
		}

		if "SEQUENCE" == comps[0] {
			sequence, err = strconv.Atoi(comps[1])
			if err != nil {
				fmt.Println("Error at reading sequence", err)
				return "", 0, err
			}
		}
	}

	return pattern, sequence, nil
}

// write to constent file
func writeVariablesToConstent(sequence int) error {
	fmt.Println("Entered in write function!")
	input, err := os.ReadFile("Constant.txt")
	if err != nil {
		fmt.Println("Error reading constant file")
		return err
	}

	lines := strings.Split(string(input), "\n")

	for i, line := range lines {
		if strings.Contains(line, "SEQUENCE") {
			str := fmt.Sprintf("SEQUENCE=%s", strconv.Itoa(sequence))
			lines[i] = str
		}
	}

	output := strings.Join(lines, "\n")
	err = os.WriteFile("Constant.txt", []byte(output), 0644)
	if err != nil {
		fmt.Println("Error at writing to the constant file")
		return err
	}

	return nil
}

func renameFile(dirPath, pattern string, sequence int) error {
	// Get all the file names in the directory
	filenames, err := filepath.Glob(filepath.Join(dirPath, "*"))
	if err != nil {
		fmt.Println("Error getting file names:", err)
		return err
	}

	// Initialize a slice to store file details for the excel sheet
	var fileDetails [][]string

	for _, filename := range filenames {
		// Generate new file name with incrementing sequence
		sequence = sequence + 1
		dir, file := filepath.Split(filename)
		var newName = fmt.Sprintf("%s%d%s", pattern, sequence, filepath.Ext(filename))
		newFileName := filepath.Join(dir, newName)

		// Rename the file
		err := changeFileName(filename, newFileName)
		if err != nil {
			fmt.Println("Error Renaming file:", err)
			return err
		}

		// Append file details to the slice
		fileDetails = append(fileDetails, []string{file, newName, filename, newFileName})
	}

	// Write the update variable back into the Constants
	err = writeVariablesToConstent(sequence)
	if err != nil {
		fmt.Println("Error writing sequence to constant variable:", err)
	}

	// Create an Excel sheet with file details
	err = createExecelSheet(dirPath, fileDetails)
	if err != nil {
		fmt.Println("Error creating Excel Sheet", err)
		return err
	}

	return nil
}

func revertFile(dirPath, excelPath string) error {
	xlsx, err := excelize.OpenFile(excelPath)
	if err != nil {
		fmt.Println("Error at opening excel file:", err)
		return err
	}

	// Get all the rows from the "Sheet 1" of the excel file
	rows, err := xlsx.Rows("Sheet1")
	if err != nil {
		fmt.Println("Error at reading rows of excel file:", err)
		return err
	}

	for rows.Next() {
		var oldFileName, newFileName string
		colsValue := rows.Columns()
		oldFileName = colsValue[0]
		newFileName = colsValue[1]

		fmt.Printf("Old: %s \t\tNew: %s", oldFileName, newFileName)
		fmt.Println()

		err := changeFileName(filepath.Join(dirPath, newFileName), filepath.Join(dirPath, oldFileName))
		if err != nil {
			fmt.Println("Error on changing name for file " + oldFileName + " to " + newFileName + ": ")
			fmt.Println(err)
			return err
		}
	}

	return nil
}

func main() {

	// Command line flags
	dirPathPtr := flag.String("dir", "foo", "Use to provide Folder's Path in which files to be rename are located.")
	excelPathPtr := flag.String("xls", "foo", "Use to provide complete path of the excel file for reverting file name")
	revertPtr := flag.Bool("r", false, "Flag use to indicate revert files to original names")

	// Parse to execute command-line parsing
	flag.Parse()

	if "foo" == *dirPathPtr {
		fmt.Println("EMPTY FOLDER PATH. you must provide folder's path.\n Exiting...")
		return
	}

	if *revertPtr && "foo" == *excelPathPtr {
		fmt.Println("EMPTY EXCEL PATH. you must provide excel file path for reverting the file name to original name\n Exiting...")
		return
	}

	if *revertPtr {
		fmt.Printf("You have entered to revert filenames of files present in %s folder.\n With Mapping present in %s.\n", *dirPathPtr, *excelPathPtr)
		fmt.Println("Please confirm the details (Enter n to cancel, any other charecter to proceed)")
		var input string
		_, err := fmt.Scan(&input)
		if err != nil {
			fmt.Println("Error:", err)
			return
		}

		if "n" == input {
			fmt.Println("You have entered \"n\" to exit\n Exiting...")
			return
		}

		fmt.Println("Starting Reverting Filenames in " + *dirPathPtr + " From the excel file at " + *excelPathPtr)
		revertFile(*dirPathPtr, *excelPathPtr)
	} else {
		pattern, sequence, err := readVariablesFromConstent()
		if err != nil {
			fmt.Println("Error in reading variable from constants:", err)
			return
		}

		fmt.Println("You have entered to rename files in folder:" + *dirPathPtr)
		fmt.Println("PATTERN in constant:", pattern)
		fmt.Println("SEQUENCE in constant:", sequence)

		fmt.Println("Please confirm the details (Enter n to cancel, any other charecter to proceed)")
		var input string
		_, err = fmt.Scan(&input)
		if err != nil {
			fmt.Println("Error:", err)
			return
		}

		if "n" == input {
			fmt.Println("You have entered \"n\" to exit\n Exiting...")
			return
		}

		fmt.Println("Starting Renaming Filenames in ", *dirPathPtr)
		err = renameFile(*dirPathPtr, pattern, sequence)
		if err != nil {
			fmt.Println("Error encounter while performing rename operation", err)
			return
		}

		fmt.Println("File names are changed succesfully. Excel Sheet created.")
	}

}
