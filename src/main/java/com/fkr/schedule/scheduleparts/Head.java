package com.fkr.schedule.scheduleparts;

import com.fkr.schedule.beautify.Styles;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class Head {

    public static void addHead(Sheet sheet, String workName, int maxTerm, int isLift) {

        // Определяем кол-во недель
        int maxWeeks;
        if (workName.equals("ТО")) maxWeeks = maxTerm;
        else maxWeeks = Math.round(maxTerm / 7);

        // Добавляем стили
        CellStyle titleBoldStyle = Styles.createTitleBoldStyle(sheet.getWorkbook());
        CellStyle titleStyle = Styles.createTitleStyle(sheet.getWorkbook());
        CellStyle tableStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.CENTER,
                false,
                false,
                (short) 0
        );

        // Создаем строки для заголовков
        Row titleBoldRow = sheet.createRow(0);
        titleBoldRow.setHeightInPoints(18f);
        Row title1 = sheet.createRow(1);
        title1.setHeightInPoints(16f);
        Row title2 = sheet.createRow(2);
        title2.setHeightInPoints(16f);

        // Заполняем заголовки
        Cell titleCellBold = titleBoldRow.createCell(0);
        titleCellBold.setCellValue("ГРАФИК ВЫПОЛНЕНИЯ РАБОТ (УСЛУГ)");
        titleCellBold.setCellStyle(titleBoldStyle);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,2+maxWeeks+isLift));

        Cell titleCell1 = title1.createCell(0);
        titleCell1.setCellValue("Начало выполнения работ: с момента подписания акта передачи для выполнения работ.");
        titleCell1.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(1,1,0,2+maxWeeks+isLift));

        Cell titleCell2 = title2.createCell(0);

        if (workName.equals("ТО")) {
            titleCell2.setCellValue("Срок окончания выполнения работ: через " + Math.round(maxWeeks/7) + " недели / "
                + maxTerm + " календарных дней с момента начала выполнения работ");
        } else {
            titleCell2.setCellValue("Срок окончания выполнения работ: через " + maxWeeks + " недель / "
                + maxTerm + " календарных дней с момента начала выполнения работ");
        }

        titleCell2.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(2,2,0,2+maxWeeks+isLift));

        // Создаем строки для шапки
        Row tableHead1 = sheet.createRow(3);
        Row tableHead2 = sheet.createRow(4);
        Row tableHead3 = sheet.createRow(5);
        Row tableHead4 = sheet.createRow(6);

        // Высота столбцов
        float tableHeadRowsHeight = 27.75f;
        tableHead1.setHeightInPoints(tableHeadRowsHeight);
        tableHead2.setHeightInPoints(tableHeadRowsHeight);
        tableHead3.setHeightInPoints(tableHeadRowsHeight);

        // Первый столбец шапки
        Cell number1 = tableHead1.createCell(0);
        Cell number2 = tableHead2.createCell(0);
        Cell number3 = tableHead3.createCell(0);
        number1.setCellValue("№ п.п объекта / вида работ / этапа");
        sheet.addMergedRegion(new CellRangeAddress(3,5,0,0));
        number1.setCellStyle(tableStyle);
        number2.setCellStyle(tableStyle);
        number3.setCellStyle(tableStyle);

        // Второй столбец шапки
        Cell nameOfWork1 = tableHead1.createCell(1);
        Cell nameOfWork2 = tableHead2.createCell(1);
        Cell nameOfWork3 = tableHead3.createCell(1);
        nameOfWork1.setCellValue("Наименование \n(объект (адрес), вид работ, технологические этапы)");
        sheet.setColumnWidth(1,14000);
        sheet.addMergedRegion(new CellRangeAddress(3, 5, 1, 1));
        nameOfWork1.setCellStyle(tableStyle);
        nameOfWork2.setCellStyle(tableStyle);
        nameOfWork3.setCellStyle(tableStyle);

        // Столбец для рег.№№, если лифты
        if (isLift == 1) {

            Cell regNum1 = tableHead1.createCell(2);
            Cell regNum2 = tableHead2.createCell(2);
            Cell regNum3 = tableHead3.createCell(2);
            regNum1.setCellValue("Регистрационный № лифта");
            regNum1.setCellStyle(tableStyle);
            regNum2.setCellStyle(tableStyle);
            regNum3.setCellStyle(tableStyle);
            sheet.setColumnWidth(2, 3500);
            sheet.addMergedRegion(new CellRangeAddress(3, 5, 2, 2));

        }

        // Третий столбец шапки (стоимость)
        Cell cost1 = tableHead1.createCell(2+isLift);
        Cell cost2 = tableHead2.createCell(2+isLift);
        Cell cost3 = tableHead3.createCell(2+isLift);
        cost1.setCellValue("Стоимость выполнения работ, отдельных видов работ, технологических этапов (руб.)");
        cost1.setCellStyle(tableStyle);
        cost2.setCellStyle(tableStyle);
        cost3.setCellStyle(tableStyle);
        sheet.setColumnWidth(2+isLift,4400);
        sheet.addMergedRegion(new CellRangeAddress(3, 5, 2+isLift, 2+isLift));

        // График работ
        for (int i=0; i<maxWeeks; i++) {

            Cell schedule = tableHead1.createCell(3+i+isLift);
            schedule.setCellStyle(tableStyle);

            Cell weeks =tableHead2.createCell(3+i+isLift);
            weeks.setCellStyle(tableStyle);

            Cell weeksNumbers = tableHead3.createCell(3+i+isLift);

            if (workName.equals("ТО")) {
                if (i==0) {
                    weeksNumbers.setCellValue(1);
                } else if (i==7) {
                    weeksNumbers.setCellValue(2);
                }
            } else {
                weeksNumbers.setCellValue(i + 1);
            }

            weeksNumbers.setCellStyle(tableStyle);

            sheet.setColumnWidth(3+i+isLift, 1120);
        }

        tableHead1.getCell(3 + isLift).setCellValue("График работ (недели)");
        tableHead2.getCell(3 + isLift).setCellValue("недели");
        sheet.addMergedRegion(new CellRangeAddress(3,3,3+isLift, maxWeeks+2+isLift));
        sheet.addMergedRegion(new CellRangeAddress(4,4,3+isLift, maxWeeks+2+isLift));

        if (workName.equals("ТО")) {
            sheet.addMergedRegion(new CellRangeAddress(5,5,3+isLift,3+isLift+6));
            sheet.addMergedRegion(new CellRangeAddress(5,5,10+isLift,10+isLift+6));
        }

        // Нумерация столбцов шапки
        for (int i=0; i<=2+maxWeeks+isLift; i++) {

            Cell tableHeadNumbering = tableHead4.createCell(i);
            tableHeadNumbering.setCellValue(i+1);
            tableHeadNumbering.setCellStyle(tableStyle);

        }

    }
}
