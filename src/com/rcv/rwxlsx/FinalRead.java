package com.rcv.rwxlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class FinalRead {

    public static void main(String[] args) throws EncryptedDocumentException, IOException, ParseException {
        File inputFile = new File("D:\\eclipse\\ExcelAnalyzer\\src\\com\\rcv\\data\\Assignment_Timecard.xlsx");
        FileInputStream fis = new FileInputStream(inputFile);

        Workbook wb = WorkbookFactory.create(fis);
        Sheet sheet1 = wb.getSheetAt(0);

        List<EmployeeShift> shifts = analyzeShifts(sheet1);

        // Print the results to the console
        for (EmployeeShift shift : shifts) {
            System.out.println("Name: " + shift.getName() + ", Position: " + shift.getPosition());
//            System.out.println("Start Time: " + shift.getStartTime() + ", End Time: " + shift.getEndTime());
//            System.out.println("==========================================");
        }

        fis.close();
        wb.close();
    }

    private static List<EmployeeShift> analyzeShifts(Sheet sheet) throws ParseException {
        List<EmployeeShift> output = new ArrayList<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm a");

        int consecutiveDays = 7;
        double minShiftGap = 1; // in hours
        double maxShiftDuration = 14; // in hours

        Row headerRow = sheet.getRow(0);

        if (headerRow == null) {
            throw new RuntimeException("Header row is null. Please check the Excel sheet format.");
        }

        int nameIndex = getCellIndex(headerRow, "Employee Name");
        int positionIndex = getCellIndex(headerRow, "Position ID");
        int timeInIndex = getCellIndex(headerRow, "Time");
        int timeOutIndex = getCellIndex(headerRow, "Time Out");

        List<EmployeeShift> allShifts = new ArrayList<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row currentRow = sheet.getRow(i);

            // Ensure current row is not null
            if (currentRow == null) {
                continue; // Skip to the next iteration if the row is null
            }

            String name = getCellValue(currentRow, nameIndex);
            String position = getCellValue(currentRow, positionIndex);

            // Additional checks for non-empty date strings
            String timeInStr = getCellValue(currentRow, timeInIndex);
            String timeOutStr = getCellValue(currentRow, timeOutIndex);
            if (timeInStr.isEmpty() || timeOutStr.isEmpty()) {
                // Skip to the next iteration if either Time In or Time Out is empty
                continue;
            }

            Date startTime = dateFormat.parse(timeInStr);
            Date endTime = dateFormat.parse(timeOutStr);

            allShifts.add(new EmployeeShift(name, position, startTime, endTime));
        }

        // Logic for consecutive days, shift gap, and shift duration
        // Example condition: If the shift duration is greater than 14 hours
        for (int i = 0; i < allShifts.size(); i++) {
            EmployeeShift currentShift = allShifts.get(i);
            Date currentEndTime = currentShift.getEndTime();

            // Check for 7 consecutive days
            if (i + consecutiveDays <= allShifts.size()) {
                boolean consecutive = true;
                for (int j = i + 1; j <= i + consecutiveDays - 1; j++) {
                    if (currentEndTime.getTime() != allShifts.get(j).getStartTime().getTime() + 1) {
                        consecutive = false;
                        break;
                    }
                    currentEndTime = allShifts.get(j).getEndTime();
                }

                if (consecutive) {
                    output.add(currentShift);
                    continue;
                }
            }

            // Check for less than 10 hours gap but greater than 1 hour
            if (i + 1 < allShifts.size()) {
                Date nextStartTime = allShifts.get(i + 1).getStartTime();
                double gapHours = (nextStartTime.getTime() - currentEndTime.getTime()) / (60 * 60 * 1000.0);
                if (1 < gapHours && gapHours < 10) {
                    output.add(currentShift);
                    continue;
                }
            }

            // Check for more than 14 hours in a single shift
            double shiftDuration = (currentShift.getEndTime().getTime() - currentShift.getStartTime().getTime())
                    / (60 * 60 * 1000.0);
            if (shiftDuration > maxShiftDuration) {
                output.add(currentShift);
            }
        }

        return output;
    }

    private static String getCellValue(Row row, int cellIndex) {
        if (cellIndex != -1) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case STRING:
                        return cell.getStringCellValue();
                    case NUMERIC:
                        // Handle numeric value
                        if (DateUtil.isCellDateFormatted(cell)) {
                            // Handle Date type
                            Date dateCellValue = cell.getDateCellValue();
                            SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm a");
                            return dateFormat.format(dateCellValue);
                        } else {
                            // Handle Numeric type
                            double numericValue = cell.getNumericCellValue();
                            // Check if the numeric value is a whole number
                            if (numericValue % 1 == 0) {
                                // If it's a whole number, convert to an integer
                                return String.valueOf((int) numericValue);
                            } else {
                                // If it's a decimal, keep it as a decimal
                                return String.valueOf(numericValue);
                            }
                        }
                    default:
                        // Handle other cell types if needed
                        return "";
                }
            }
        }
        return "";
    }

    private static int getCellIndex(Row headerRow, String columnName) {
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return i;
            }
        }
        // Return a value indicating that the cell was not found
        return -1;
    }

    private static class EmployeeShift {
        private final String name;
        private final String position;
        private final Date startTime;
        private final Date endTime;

        public EmployeeShift(String name, String position, Date startTime, Date endTime) {
            this.name = name;
            this.position = position;
            this.startTime = startTime;
            this.endTime = endTime;
        }

        public String getName() {
            return name;
        }

        public String getPosition() {
            return position;
        }

        public Date getStartTime() {
            return startTime;
        }

        public Date getEndTime() {
            return endTime;
        }
    }
}
