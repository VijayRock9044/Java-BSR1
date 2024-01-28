package com.scb.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelValidatorService {

    public String validate() {
        String message = "";
        try {
            String directory = "C:\\Users\\Vijay\\Documents\\Automation_Code\\Files";

            // Step 1: Read data from multiple Excel files
            List<String> excelFiles = getExcelFilesFromSharedPath(directory + "\\IN");
            List<List<String>> allData = new ArrayList<>();

            for (String file : excelFiles) {
                List<List<String>> data = readExcelFile(directory + "\\IN\\" + file);
                allData.addAll(data);
            }

            // Step 2: Collate all data into one Excel file
            writeToExcel(allData, "collatedData.xlsx");

            List<List<String>> masterdata = readExcelFile(directory + "\\Masterdata.xlsx");

            // Verify the data against masterdata
            List<String> invalidMasterRows = validateMasterData(allData, masterdata);

            // Step 3: Validate the data
            List<String> invalidRows = validateData(allData);

            invalidRows.addAll(invalidMasterRows);

            if (invalidRows.isEmpty()) {
                // Step 4: Export error-free data to a new Excel file
                writeToExcel(allData, directory + "\\OUT\\errorFreeData.xlsx");
                writeToText(allData, directory + "\\OUT\\errorFreeData.txt");
                message = "Data validation successful. Error-free file exported.";
            } else {
                message = "Data validation failed. Invalid rows: " + invalidRows;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return message;
    }

    private static List<String> getExcelFilesFromSharedPath(String sharedPath) {
        // Implement logic to get a list of Excel files from the shared path
        // Example: List all files and filter based on the extension (.xlsx)
        List<String> fileNames = new ArrayList<>();
        File folder = new File(sharedPath);
        for (File file: folder.listFiles()) {
            fileNames.add(file.getName());
        }
        return fileNames;
    }

    private static List<List<String>> readExcelFile(String filePath) throws IOException {
        List<List<String>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            int count = 1;

            for (Row row : sheet) {
                if (count == 1) {
                    count++;
                    continue;
                }
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    cell.setCellType(CellType.STRING);
                    rowData.add(cell.toString());
                }
                data.add(rowData);
                count++;
            }
        }

        return data;
    }

    private static List<String> validateMasterData(List<List<String>> data, List<List<String>> masterData) {
        // Implement validation logic based on the given validation rules
        // Store the row numbers with validation errors in a list
        List<String> invalidRows = new ArrayList<>();

        // Example validation (adjust as needed)
        for (int i = 0; i < data.size(); i++) {
            List<String> rowData = data.get(i);

            if (rowData.size() != 23) {
                invalidRows.add("The row " + String.valueOf(i + 1) + " does not have all required 23 columns"); // Excel rows are 1-indexed
            }

            // Implement additional validation logic for each column based on the specified rules
            for (int j = 0; j < rowData.size(); j++) {
                int expectedIterations = getExpectedIterations(j);
                if (expectedIterations == -1) {
                    int findCount = 0;
                    while (findCount != -1) {
                        try {
                            masterData.get(findCount).get(j);
                            findCount++;
                        } catch (IndexOutOfBoundsException iobe) {
                            expectedIterations = findCount;
                            findCount = -1;
                        } catch (Exception e) {
                            e.printStackTrace();
                            expectedIterations = findCount;
                            findCount = -1;
                        }
                    }
                }
                boolean valid = false;
                String cellValue = rowData.get(j);
                for (int k = 0; k < expectedIterations; k++) {
                    if (i == 17 && j ==13) {
                        int e = 0;
                    }
                    if (cellValue.equalsIgnoreCase(masterData.get(k).get(j))) {
                        valid = true;
                        break;
                    }
                }
                if (!valid && expectedIterations > 0) {
                    System.out.println("Data " + cellValue + " at Row " + i + " and Column " + j + " is not present in master data");
                    invalidRows.add("Data " + cellValue + " at Row " + i + " and Column " + j + " is not present in master data");
                }
            }
        }

        return invalidRows;
    }

    private static List<String> validateData(List<List<String>> data) {
        // Implement validation logic based on the given validation rules
        // Store the row numbers with validation errors in a list
        List<String> invalidRows = new ArrayList<>();

        // Example validation (adjust as needed)
        for (int i = 0; i < data.size(); i++) {
            List<String> rowData = data.get(i);

            if (rowData.size() != 23) {
                invalidRows.add("The row " + String.valueOf(i + 1) + " does not have all required 23 columns"); // Excel rows are 1-indexed
            }

            // Implement additional validation logic for each column based on the specified rules
            // Example: Check the length of each column
            for (int j = 0; j < rowData.size(); j++) {
                int expectedLength = getExpectedLength(j);
                if (expectedLength == 10 && rowData.get(j).length() > 10) {
                    System.out.println("Row " + j + " with data " + rowData.get(j) + " does not have valid data");
                    invalidRows.add("Row " + j + " with data " + rowData.get(j) + " does not have valid data");
                } else if (expectedLength != 10 && rowData.get(j).length() != expectedLength) {
                    System.out.println("Row " + j + " with data " + rowData.get(j) + " does not have valid data");
                    invalidRows.add("Row " + j + " with data " + rowData.get(j) + " does not have valid data");
                } else {
//                    System.out.println("Row " + j + " with data " + rowData.get(j) + " has valid data");
                }
            }
        }

        return invalidRows;
    }

    private static int getExpectedLength(int columnIndex) {
        // Implement logic to return the expected length based on the specified validation rules
        // You can use a switch statement or another method to handle different column lengths
        // Example: Return the expected length for each column
        return switch (columnIndex) {
            case 0, 1, 11, 12, 14 -> 2;
            case 2, 6, 7, 18, 21 -> 4;
            case 3, 5 -> 0;
            case 4 -> 7;
            case 8, 10, 15, 16, 17 -> 1;
            case 9 -> 3;
            case 13 -> 5;
            case 22 -> 22;
            case 19, 20 -> 10;
            default -> -1; // Handle unknown columns
        };
    }

    private static int getExpectedIterations(int columnIndex) {
        // Implement logic to return the expected iterations based on the specified validation rules
        // You can use a switch statement or another method to handle different column lengths
        return switch (columnIndex) {
            case 0, 1, 2, 6, 7, 21 -> 1;
            case 3, 5, 18, 19, 20, 22 -> 0;
            default -> -1; // Handle unknown columns
        };
    }

    private static void writeToExcel(List<List<String>> data, String fileName) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(fileName)) {

            Sheet sheet = workbook.createSheet("Data");

            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                List<String> rowData = data.get(i);

                for (int j = 0; j < rowData.size(); j++) {
                    String cellValue = rowData.get(j);
                    if (j == 19 || j == 20) {
                        int zeroLength = 10 - cellValue.length();
                        String zeros = "";
                        for (int k = 0; k < zeroLength; k++) {
                            zeros += "0";
                        }
                        cellValue =  zeros + cellValue;
                    }
                    Cell cell = row.createCell(j);
                    cell.setCellValue(cellValue);
                }
            }

            workbook.write(fos);
        }
    }

    private static void writeToText(List<List<String>> data, String fileName) throws IOException {
        try (PrintWriter fos = new PrintWriter(fileName)) {

            for (int i = 0; i < data.size(); i++) {
                List<String> rowData = data.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    String cellValue = rowData.get(j);
                    if (j == 19 || j == 20) {
                        int zeroLength = 10 - cellValue.length();
                        String zeros = "";
                        for (int k = 0; k < zeroLength; k++) {
                            zeros += "0";
                        }
                        cellValue =  zeros + cellValue;
                    }
                    if (cellValue.isEmpty()) {
                        cellValue = " ";
                    }
                    fos.print(cellValue);
                }
                fos.println("");
            }
        }
    }
}
