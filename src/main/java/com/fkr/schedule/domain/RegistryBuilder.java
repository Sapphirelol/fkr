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

        WorkNames workNames = WorkNamesBuilder.buildFromWorkNamesFile("F:/Schedules/Виды работ.xlsx");

        Row titleRow = sheet.getRow(6);
        int addressColumnIndex = 0,
                statusColumnIndex = 0,
                regNumColumnIndex = 0,
                workNameColumnIndex = 0,
                termsColumnIndex = 0,
                costColumnIndex = 0,
                typeColumnIndex = 0;

        for (int j=0; j<20; j++) {

            try {

                if (titleRow.getCell(j).getStringCellValue().isEmpty() || titleRow.getCell(j).getStringCellValue().equals("")) {
                    System.out.println("Конец шапки таблицы, всего колонок - " + (j+1));
                    break;
                } else if (addressColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Адрес")) {
                    addressColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (statusColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("КГИОП")) {
                    statusColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (regNumColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("лифт")) {
                    regNumColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (workNameColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Вид работ")) {
                    workNameColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (termsColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Срок выполнения работ")) {
                    termsColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (costColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Сметная стоимость")) {
                    costColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                } else if (typeColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Тип") & !titleRow.getCell(j).getStringCellValue().contains("Тип шахты")) {
                    typeColumnIndex = j;
                    System.out.println(titleRow.getCell(j).getStringCellValue());
                }

            } catch (Exception e) {
                System.out.println("Ошибка в шапке реестра");
                break;
            }

        }

        for (int i=7; i<=sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            try {
                if (row.getCell(addressColumnIndex).getStringCellValue().isEmpty() ||
                    row.getCell(addressColumnIndex).getStringCellValue().equals(""))
                {
                    break;
                }
            } catch (Exception e) {
                break;
            }

            registry.getAddresses().add(row.getCell(addressColumnIndex).getStringCellValue());
            registry.getStatus().add(row.getCell(statusColumnIndex).getStringCellValue().equals("ОКН"));
            registry.getRegNum().add(row.getCell(regNumColumnIndex).getStringCellValue());

            int indexOfWorkName = workNames.getLongNames().indexOf(row.getCell(workNameColumnIndex).getStringCellValue());

            registry.getWorkNames().add(workNames.getShortNames().get(indexOfWorkName));

            if (
                    typeColumnIndex != 0 &
                    !row.getCell(typeColumnIndex).getStringCellValue().isEmpty() &
                    !row.getCell(typeColumnIndex).getStringCellValue().equals("")
            ) {
                registry.getWorkTypes().add(" " + row.getCell(typeColumnIndex).getStringCellValue());
            } else {
                registry.getWorkTypes().add("");
            }

            try {
                registry.getTerms().add(Integer.valueOf(row.getCell(termsColumnIndex).getStringCellValue()));
            } catch (Exception ignored) {
                try {
                    registry.getTerms().add((int) Math.rint(row.getCell(termsColumnIndex).getNumericCellValue()));
                } catch (Exception e) {
                    System.out.println("Ошибка в поле \"Срок\" в строке " + i);
                }
            }

            registry.getCost().add(row.getCell(costColumnIndex).getNumericCellValue());

        }

        registry.setMaxTerm(Collections.max(registry.getTerms()));

        return registry;
    }
}
