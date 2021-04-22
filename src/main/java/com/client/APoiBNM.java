package com.client;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class APoiBNM {
    public static void main(String[] args) throws Exception {
        createWorkBook();
        List<ValCurs> valCurs = new ArrayList<>();
        List<String> listOfDates = ReadCSVDates.parserCSV();

        for (String str : listOfDates) {
            valCurs.add(BNMGet.sendGet(str));
        }

        for (ValCurs temp : valCurs) {
            openWorkBook(temp);
        }
    }

    private static void createWorkBook() throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        FileOutputStream out = new FileOutputStream("ExcelValutes.xlsx");

        workbook.write(out);
        out.close();
    }


    private static void openWorkBook(ValCurs valCurs) throws Exception {
        File file = new File("ExcelValutes.xlsx");
        FileInputStream fIP = new FileInputStream(file);

        XSSFWorkbook workbook = new XSSFWorkbook(fIP);
        XSSFSheet spreadsheet = workbook.createSheet(valCurs.getDate());

        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        XSSFRow row;

        Map<String, Object[]> valutes = new TreeMap<>();

        valutes.put("*", new Object[]{"Name", "Date"});

        valutes.put("+", new Object[]{valCurs.getName(), valCurs.getDate()});

        valutes.put("/", new Object[]{"ID", "NumCode", "CharCode", "Nominal",
                "Name", "Value"});

        for (Valute valute : valCurs.getList()) {
            valutes.put(valute.getId(), new Object[]{
                    valute.getId(),
                    valute.getNumCode(),
                    valute.getСharCode(),
                    valute.getNominal() + "",
                    valute.getName(),
                    Double.toString(valute.getValue())});
        }

        Set<String> keyid = valutes.keySet();
        int rowid = 0;

        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);

            Object[] objectArr = valutes.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                spreadsheet.autoSizeColumn(cellid);

                if (rowid == 1 || rowid == 3) {
                    cell.setCellStyle(style);
                }
                cell.setCellValue((String) obj);
            }
        }

        FileOutputStream out = new FileOutputStream("ExcelValutes.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("Excel file written successfully");
    }
}