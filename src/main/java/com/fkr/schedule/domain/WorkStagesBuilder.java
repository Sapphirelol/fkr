package com.fkr.schedule.domain;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public abstract class WorkStagesBuilder {
    public static WorkStages buildFromWorkStagesFile(String file) throws IOException {

        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheetAt(0);

        WorkStages workStages = new WorkStages(
                sheet.getRow(0).getCell(1).getStringCellValue(),
                new ArrayList<String>(),
                new ArrayList<Short>(),
                new ArrayList<Short>()
        );

        for (int i=0; i<sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i+1);

            workStages.getNames().add(row.getCell(1).getStringCellValue());
            workStages.getWeeksTo().add((short) Math.round(row.getCell(2).getNumericCellValue()));
            workStages.getWeeksFor().add((short) Math.round(row.getCell(3).getNumericCellValue()));

        }

        return workStages;
    }
}
