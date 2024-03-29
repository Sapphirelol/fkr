package com.fkr.schedule.scheduleparts;

import com.fkr.schedule.beautify.Styles;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

class LiftStages {

    static void addLiftStages(
            Sheet sheet,
            int weeksFull,
            String workName,
            int addressNum,
            int workNum,
            int designTerm
    ) {

        // Стили
        CellStyle tableStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.CENTER,
                false,
                false,
                (short) 0
        );

        CellStyle nameStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.LEFT,
                false,
                false,
                (short) 0
        );

        Workbook wb = new XSSFWorkbook();

        try {
            if (workName.contains("Лифт")) {
                wb = new XSSFWorkbook(new FileInputStream(
                        "Этапы/Лифты/" +
                                workName +
                                " " +
                                weeksFull * 7 +
                                ".xlsx")
                );
            } else {
                wb = new XSSFWorkbook(new FileInputStream(
                        "Этапы/Лифты/" +
                                workName +
                                " " +
                                weeksFull +
                                ".xlsx")
                );

            }
        } catch (Exception e) {
            System.out.println("Ошибка в чтении шаблона с этапами");
            System.out.println("workName - " + workName + " weeksFull - " + weeksFull);
        }

        Sheet templateSheet = wb.getSheetAt(0);

        int stageCount = workName.equals("ТО") || workName.equals("ЛифтПД") ? 3 : 8;

        for (int k = 0; k < stageCount; k++) {

            // Строка под этап
            Row stageRow = sheet.createRow(sheet.getLastRowNum() + 1);

            // Номер этапа
            Cell stageNumCell = stageRow.createCell(0);
            stageNumCell.setCellValue(addressNum + "." + workNum + "." + (k + 1));
            stageNumCell.setCellStyle(tableStyle);

            // Наименование этапа
            Cell stageNameCell = stageRow.createCell(1);
            stageNameCell.setCellValue(templateSheet.getRow(k).getCell(1).getStringCellValue());
            stageNameCell.setCellStyle(nameStyle);

            // Рег.№
            Cell stageRegNumCell = stageRow.createCell(2);
            stageRegNumCell.setCellStyle(tableStyle);

            // Сумма этапа (-)
            Cell stageSumCell = stageRow.createCell(3);
            stageSumCell.setCellStyle(tableStyle);


            // Заполнение этапа
            for (int j=designTerm; j < sheet.getRow(6).getLastCellNum() - 4; j++) {

                Cell stageCell = stageRow.createCell(4 + j);
                stageCell.setCellStyle(tableStyle);

                if (j < weeksFull) {

                    double percent = templateSheet.getRow(k).getCell(2 + j).getNumericCellValue();

                    if (percent != 0d) {
                        stageCell.setCellValue(templateSheet.getRow(k).getCell(2 + j).getNumericCellValue());
                    }

                }

            }

        }

    }
}
