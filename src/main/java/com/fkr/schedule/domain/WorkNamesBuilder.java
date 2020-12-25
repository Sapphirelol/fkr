package com.fkr.schedule.domain;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public abstract class WorkNamesBuilder {

    public static WorkNames buildFromWorkNamesFile(String file) throws IOException {

        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheet("Виды работ");

        WorkNames workNames = new WorkNames(
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>()
        );

        for (int i=0; i<=sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            workNames.getIds().add((int) row.getCell(0).getNumericCellValue());
            workNames.getShortNames().add(row.getCell(1).getStringCellValue());
            workNames.getLongNames().add(row.getCell(2).getStringCellValue());

        }

        return workNames;
    }

}
