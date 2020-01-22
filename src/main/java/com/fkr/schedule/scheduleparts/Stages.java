package com.fkr.schedule.scheduleparts;

import com.fkr.schedule.beautify.Styles;
import com.fkr.schedule.domain.WorkStages;
import org.apache.poi.ss.usermodel.*;

public class Stages {

    public static void addStages(
            Sheet sheet,
            WorkStages workStages,
            int weeksFull,
            int weeksForPrep,
            int addressNum,
            int workNum
    ) {

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


        int weeksBuffer = 0;

        for (int stageCount=0; stageCount < workStages.getNames().size(); stageCount++) {

            // Строка под этап
            Row stageRow = sheet.createRow(sheet.getLastRowNum() + 1);

            // Номер этапа
            Cell stageNumCell = stageRow.createCell(0);
            stageNumCell.setCellValue(addressNum + "." + workNum + "." + (stageCount + 1));
            stageNumCell.setCellStyle(tableStyle);

            // Наименование этапа
            Cell stageNameCell = stageRow.createCell(1);
            stageNameCell.setCellValue(workStages.getNames().get(stageCount));
            stageNameCell.setCellStyle(nameStyle);

            // Сумма этапа (-)
            Cell stageSumCell = stageRow.createCell(2);
            stageSumCell.setCellStyle(tableStyle);


            // График выполнения этапа
            int weeksTo = 0;
            int weeksFor = 1;
            float coefficient = 0.0f;

            if (workStages.getWorkName().equals("Ремонт крыши (ж)")) {

                if (stageCount == 0) {

                    weeksFor = weeksForPrep - 1;

                } else if (stageCount == 1) {

                    weeksTo = weeksForPrep - 1;

                } else {

                    coefficient = (weeksFull-weeksForPrep) / 10f;
                    weeksTo = Math.round(workStages.getWeeksTo().get(stageCount) * coefficient) + weeksForPrep;
                    weeksFor = Math.max(Math.round(workStages.getWeeksFor().get(stageCount) * coefficient), 1);

                    if ((weeksFor+weeksTo)>weeksFull) {

                        weeksTo = weeksFull - 1;
                        weeksFor = 1;

                    }

                }

            } else {

                int x = 1;
                while (weeksForPrep >= x) {

                    Cell cell = stageRow.createCell(2 + x++);
                    cell.setCellStyle(tableStyle);

                }

                coefficient = (weeksFull - weeksForPrep) / 13.0f;
                weeksTo = Math.abs(Math.round(workStages.getWeeksTo().get(stageCount) * coefficient));
                weeksFor = Math.max(Math.round(workStages.getWeeksFor().get(stageCount) * coefficient), 1);

                if ((weeksFor+weeksTo)>(weeksFull-weeksForPrep)) {

                    weeksTo = weeksFull - weeksForPrep - 1;
                    weeksFor = 1;

                }

            }

            System.out.println(workStages.getNames().get(stageCount));
            System.out.print(weeksTo + " | ");
            System.out.print((workStages.getWeeksTo().get(stageCount) * coefficient) + " || ");
            System.out.print(weeksFor + " | ");
            System.out.print((workStages.getWeeksFor().get(stageCount) * coefficient) + " || ");
            System.out.println(coefficient);

            // Заполнение %
            int wfp = weeksForPrep;
            if (workStages.getWorkName().equals("Ремонт крыши (ж)")) {
                weeksForPrep = 0;
            }

            int percent = (100 - 100 % weeksFor) / weeksFor;
            int b = 100 % weeksFor;
            for (int m=0; m < (weeksFull - weeksForPrep); m++) {
                Cell cell = stageRow.createCell(3 + weeksForPrep + m);
                cell.setCellStyle(tableStyle);
                if ((weeksTo <= m) && (weeksFor > m - weeksTo)) {
                    if (b > 0) {
                        cell.setCellValue(percent + 1);
                        b = b - 1;
                    } else {
                        cell.setCellValue(percent);
                    }
                }
            }

            weeksForPrep = wfp;

            for (int m=stageRow.getLastCellNum(); m < sheet.getRow(sheet.getLastRowNum()-1).getLastCellNum(); m++) {

                Cell cell = stageRow.createCell(m);
                cell.setCellStyle(tableStyle);

            }

        }

    }
}
