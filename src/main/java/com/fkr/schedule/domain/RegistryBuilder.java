package com.fkr.schedule.domain;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;

public abstract class RegistryBuilder {
    public static Registry buildFromRegistryFile(String file) throws Exception {

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

        WorkNames workNames = WorkNamesBuilder.buildFromWorkNamesFile("Виды работ.xlsx");

        Row titleRow = sheet.getRow(6);
        int addressColumnIndex = 0,
                statusColumnIndex = 0,
                regNumColumnIndex = 0,
                workNameColumnIndex = 0,
                termsColumnIndex = 0,
                costColumnIndex = 0,
                typeColumnIndex = 0;

        for (int j=0; j<titleRow.getLastCellNum(); j++) {

            try {

                if (addressColumnIndex == 0 & titleRow.getCell(j).getStringCellValue().contains("Адрес")) {
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
                System.out.println("Ошибка при чтении шапки реестра, колонка - " + (j+1));
                break;
            }

        }

        int startRowNum = 7;

        for (int i=startRowNum; i<sheet.getLastRowNum(); i++){

            Row row=sheet.getRow(i);

            try {
                if (
                        row.getCell(addressColumnIndex) == null ||
                        row.getCell(addressColumnIndex).getStringCellValue().isEmpty() ||
                        row.getCell(addressColumnIndex).getStringCellValue().equals("")
                ) {
                    System.out.println("Конец реестра");
                    break;
                }
            } catch (Exception e) {
                System.out.println("Конец реестра");
                break;
            }

            try {
                registry.getAddresses().add(row.getCell(addressColumnIndex).getStringCellValue());
            } catch (Exception e) {
                System.out.println("Ошибка в чтении адреса");
            }

            try {
                registry.getStatus().add(row.getCell(statusColumnIndex).getStringCellValue().contains("ОКН"));
            } catch (Exception e) {
                System.out.println("Ошибка в чтении статуса");
            }

            try {
                registry.getRegNum().add(row.getCell(regNumColumnIndex).getStringCellValue());
            } catch (Exception e) {
                System.out.println("Ошибка в чтении рег.№");
            }

            int indexOfWorkName=0;

            try {
                indexOfWorkName=workNames.getLongNames().indexOf(row.getCell(workNameColumnIndex).getStringCellValue());
            } catch (Exception e) {
                System.out.println("Ошибка при определении вида работ в шаблонах");
            }

            try {
                registry.getWorkNames().add(workNames.getShortNames().get(indexOfWorkName));
            } catch (Exception e) {
                System.out.println("Ошибка в чтении ...");
            }

            try {
                if (typeColumnIndex != 0) {
                    registry.getWorkTypes().add(row.getCell(typeColumnIndex).getStringCellValue());
                } else {
                    registry.getWorkTypes().add("");
                }
            } catch (Exception e) {
                System.out.println("Ошибка при добавлении подвида работ");
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

            try {
                registry.getCost().add(row.getCell(costColumnIndex).getNumericCellValue());
            } catch (Exception ignored) {
                try {
                    registry.getCost().add(Double.valueOf(row.getCell(costColumnIndex).getStringCellValue()));
                } catch (Exception e) {
                    System.out.println("Ошибка в поле \"Стоимость\" в строке " + i);
                }
            }

            registry.setDesignTerm(0);
            // Обновляем срок по проектированию на каждой итерации
            if (registry.getWorkNames().get(i - startRowNum).contains("ПД") &
                    registry.getTerms().get(i - startRowNum) > registry.getDesignTerm()) {
                registry.setDesignTerm(registry.getTerms().get(i - startRowNum));
            }

        }

        // Обновляем максимальный срок по всему графику на каждой итерации: макс.СМР + макс.ПД
        if (registry.getTerms().isEmpty()) {
            System.out.println("Нет данных для определения максимального срока");
        } else {
            registry.setMaxTerm(Collections.max(registry.getTerms()) + registry.getDesignTerm());
        }

        System.out.println("Данные из реестра получены");

        return registry;
    }
}
