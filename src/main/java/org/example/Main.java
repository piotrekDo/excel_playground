package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

/**
 * formuly
 * <p>
 * H X sumuje kolumny od 10 dosiebie minus 1
 * D7 wskazuje na HX
 */
public class Main {
    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(new File("src/main/resources/template.xlsx"));

        int daysToAdd = 31;
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        fillEmployee(sheet);
        fillMonth(sheet);

        addDayRow(sheet, daysToAdd);
        fillDates(sheet, daysToAdd);
        int summaryCell = addDaysSummary(sheet, daysToAdd);
        addTimeOnProjectsCellPointer(sheet, summaryCell);

        // Zapisz zmodyfikowany plik
        FileOutputStream outFile = new FileOutputStream(new File("output.xlsx"));
        workbook.write(outFile);

        // Zamknij strumienie
        outFile.close();
        workbook.close();
        file.close();
    }

    static void addTimeOnProjectsCellPointer(Sheet sheet, int summaryCellRow) {
        Row row = sheet.getRow(6);
        Cell cell = row.getCell(3);
        cell.setCellFormula(String.format("H%s", summaryCellRow));
    }

    static int addDaysSummary(Sheet sheet, int days) {
        Row summaryRow = sheet.createRow(10 + days);
        for (int i = 0; i < 10; i++) {
            summaryRow.createCell(i);
        }
        Cell descririptionCell = summaryRow.getCell(6);
        Cell summaryCell = summaryRow.getCell(7);
        descririptionCell.setCellValue("Suma godzin:");
        summaryCell.setCellFormula(String.format("SUM(H11:H%d)", 10 + days));
        return 10 + days;
    }

    static void fillDates(Sheet sheet, int days) {
        Row row;

        for (int i = 0; i < days; i++) {
            row = sheet.getRow(10 + i);
            Cell cell = row.getCell(2);
            LocalDate localDate = LocalDate.of(2024, 12, i + 1);
            Date date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);

            if (localDate.getDayOfWeek() == DayOfWeek.SATURDAY || localDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                for (int i1 = 0; i1 < 8; i1++) {
                    Cell cell1 = row.getCell(2 + i1);

                    // Pobierz istniejący styl
                    CellStyle originalStyle = cell1.getCellStyle();
                    Workbook workbook = sheet.getWorkbook();

                    // Skopiuj styl
                    CellStyle newStyle = workbook.createCellStyle();
                    newStyle.cloneStyleFrom(originalStyle);

                    // Dodaj szare tło
                    newStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                    newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                    // Zastosuj nowy styl do komórki
                    cell1.setCellStyle(newStyle);
                }
            }
        }
    }

    static void addDayRow(Sheet sheet, int daysToAdd) {
        Row sourceRow = sheet.getRow(10);

        for (int i = 0; i < daysToAdd - 1; i++) {
            Row newRow = sheet.createRow(11 + i);
            for (int j = 0; j < 10; j++) {
                Cell sourceCell = sourceRow.getCell(j);
                if (sourceCell != null) {
                    Cell targetCell = newRow.createCell(j);
                    targetCell.setCellStyle(sourceCell.getCellStyle());
                }
            }
        }

    }

    static void fillEmployee(Sheet sheet) {
        Row row = sheet.getRow(2);
        Cell cell = row.getCell(3);
        cell.setCellValue("Piotr Domagalski");
    }

    static void fillMonth(Sheet sheet) {
        Row row = sheet.getRow(3);
        Cell cell = row.getCell(3);
        cell.setCellValue("grudzień 2024");
    }
}