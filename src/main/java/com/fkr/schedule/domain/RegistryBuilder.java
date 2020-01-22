package com.fkr.schedule.domain;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;

public abstract class RegistryBuilder {
    public static Registry buildFromRegistryFile(String file) throws IOException {

        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheet("Реестр");

        Registry registry = new Registry(
                file,
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>(),
                new ArrayList<>()
        );

        for (int i=0; i<sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i+1);

            registry.getAddresses().add(row.getCell(1).getStringCellValue());
            registry.getStatus().add(row.getCell(2).getStringCellValue().equals("ОКН"));
            registry.getRegNum().add(row.getCell(3).getStringCellValue());
            registry.getWorkNames().add(row.getCell(4).getStringCellValue());

            if (row.getCell(4).getStringCellValue().equals("Ремонт крыш") ||
                row.getCell(4).getStringCellValue().equals("Ремонт фасадов"))
            {
                registry.getWorkTypes().add(row.getCell(5).getStringCellValue());
            } else {
                registry.getWorkTypes().add("");
            }

            registry.getTerms().add((short) Math.rint(row.getCell(7).getNumericCellValue()));
            registry.getCost().add(row.getCell(6).getNumericCellValue());

        }

        registry.setMaxTerm(Collections.max(registry.getTerms()));

        return registry;
    }
}
