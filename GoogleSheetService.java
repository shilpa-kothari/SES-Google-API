package SESGoogleSheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFSheet;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public class GoogleSheetService {
    private static final String SPREADSHEET_ID = "15etXathuHBewTOjqT05ZbqcqWrwfzFnuAq_6HCDoK10"; // ID of the target Google Sheets document
    private static final String SHEET_NAME = "Raj ERP 3.0 Tickets"; // Name of the target sheet tab

    public static void replaceSheetData(String sourceSheetId) throws IOException, GeneralSecurityException {
        Sheets sheetsService = SheetsServiceUtil.getSheetsService();

        // Parse the downloaded Excel file
        //List<List<Object>> data = parseExcelFile("storage/Book2.xls");
        // Retrieve data from the source sheet
        List<List<Object>> data = retrieveDataFromSheet(sheetsService, sourceSheetId, "RERP3_Ticket_Details");

        // Update the target sheet tab with the parsed data
        if (data != null && !data.isEmpty()) {
            updateSheetTab(sheetsService, data);
        } else {
            System.out.println("No data found in the Excel file.");
        }
        
    }
    
    
    private static List<List<Object>> retrieveDataFromSheet(Sheets service, String spreadsheetId, String sheetName) throws IOException {
        String range = "!A1:ZZ"; // Change the range as needed
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        return response.getValues();
    }
    
    private static List<List<Object>> parseExcelFileXls(String filePath) {
        List<List<Object>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             HSSFWorkbook workbook = new HSSFWorkbook(fis)) { // Use HSSFWorkbook for XLS format

            HSSFSheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            for (Row row : sheet) {
                List<Object> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    switch (cell.getCellType()) { // Use getCellTypeEnum() instead of getCellType() for HSSF
                        case STRING:
                            rowData.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            rowData.add(cell.getNumericCellValue());
                            break;
                        // Handle other cell types as needed
                        default:
                            rowData.add(null);
                            break;
                    }
                }
                data.add(rowData);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return data;
    }


    private static List<List<Object>> parseExcelFile(String filePath) {
        List<List<Object>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            for (Row row : sheet) {
                List<Object> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            rowData.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            rowData.add(cell.getNumericCellValue());
                            break;
                        // Handle other cell types as needed
                        default:
                            rowData.add(null);
                            break;
                    }
                }
                data.add(rowData);
            }
        } catch (IOException  e) {
            e.printStackTrace();
        }

        return data;
    }

    private static void updateSheetTab(Sheets sheetsService, List<List<Object>> data) throws IOException {
        // Prepare the request to update the target sheet tab
        String range = SHEET_NAME + "!A1"; // Starting cell to write data in the target sheet
        ValueRange body = new ValueRange().setValues(data);

     // Calculate the difference in row count between old and new data
        String oldRange = SHEET_NAME + "!A1:ZZ"; // Assuming the old data occupies columns A, B, and C
        int oldRowCount = getRowCount(sheetsService, SPREADSHEET_ID, oldRange);
        int newRowCount = body.size();
        int rowCountDifference = oldRowCount - newRowCount;

        System.out.printf("%d old sheet data. %d", oldRowCount, newRowCount);
        if (rowCountDifference > 0) {
            // Clear excess rows in the range
            String clearRange = SHEET_NAME + "!A" + (newRowCount + 1) + ":Z" + oldRowCount;
            ClearValuesRequest requestBody = new ClearValuesRequest();
            sheetsService.spreadsheets().values().clear(SPREADSHEET_ID, clearRange, requestBody).execute();
        }

        
        // Execute the update
        UpdateValuesResponse result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, range, body)
                .setValueInputOption("RAW")
                .execute();

        System.out.printf("%d cells updated in the target sheet.", result.getUpdatedCells());
    }
    
    private static int getRowCount(Sheets sheetsService, String spreadsheetId, String range) throws IOException {
        ValueRange response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        return values != null ? values.size() : 0;
    }
}
