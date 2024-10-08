package excelExample;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelExample {

    public static void main(String[] args) {
        String filePath = "data.xlsx";

       
        writeDataToExcel(filePath);

  
        readDataFromExcel(filePath);
    }

    private static void writeDataToExcel(String filePath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("Email");

     
        String[][] data = {
                {"John Doe", "30", "john@test.com"},
                {"Jane Doe", "28", "jane@test.com"},
                {"Bob Smith", "35", "jacky@example.com"},
                {"Swapnil", "37", "swapnil@example.com"}
        };

        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                row.createCell(j).setCellValue(data[i][j]);
            }
        }

       
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Data written to Excel file successfully.");
    }

    private static void readDataFromExcel(String filePath) {
        try (FileInputStream fileIn = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fileIn)) {
            Sheet sheet = workbook.getSheet("Sheet1");

            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            break;
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
