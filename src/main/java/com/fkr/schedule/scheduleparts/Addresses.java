package com.fkr.schedule.scheduleparts;

import com.fkr.schedule.beautify.Styles;
import com.fkr.schedule.domain.Registry;
import com.fkr.schedule.domain.WorkStages;
import com.fkr.schedule.domain.WorkStagesBuilder;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;

public class Addresses {

    public static void addAddresses(Sheet sheet, Registry registry, int isLift) throws IOException {

        int addressNum = 0;
        int workNum = 0;

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

        CellStyle addressStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.LEFT,
                true,
                false,
                (short) 0
        );

        CellStyle workStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.LEFT,
                false,
                true,
                (short) 0
        );

        CellStyle costStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.CENTER,
                false,
                false,
                (short) 4
        );


        for (int workIndex=0; workIndex < registry.getAddresses().size(); workIndex++) {

            if (workIndex ==0 || !registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex -1))) {

                // Строка под адрес
                Row addressRow = sheet.createRow(sheet.getLastRowNum() + 1);

                // Ячейка для номера адреса
                Cell addressNumCell = addressRow.createCell(0);
                addressNumCell.setCellValue(++addressNum);
                addressNumCell.setCellStyle(tableStyle);

                // Ячейка для наименования адреса
                Cell addressNameCell = addressRow.createCell(1);
                addressNameCell.setCellValue(registry.getAddresses().get(workIndex));
                addressNameCell.setCellStyle(addressStyle);

                for (int j = addressRow.getLastCellNum(); j < sheet.getRow(6).getLastCellNum(); j++) {
                    Cell cell = addressRow.createCell(j);
                    cell.setCellStyle(tableStyle);
                }

            }

            // Добавляем этапы из файла шаблона (если не лифты)
            WorkStages workStages = null;
            if (isLift == 0 || registry.getWorkNames().get(workIndex).contains("ЛО_")) {
                if (registry.getWorkTypes().get(workIndex).equals("")) {
                    workStages = WorkStagesBuilder.buildFromWorkStagesFile("Этапы/" +
                            registry.getWorkNames().get(workIndex) + ".xlsx");
                } else {
                    workStages = WorkStagesBuilder.buildFromWorkStagesFile("Этапы/" +
                            registry.getWorkNames().get(workIndex) + " " + registry.getWorkTypes().get(workIndex) + ".xlsx");
                }
            }

            // Определяем наличие этапа ОКН/ГАТИ
            int weeksForPrep = 0;
            
            // Определение недель на подготовительный этап если необходим ордер ГАТИ
            try {
                if (
                        registry.getWorkTypes().get(workIndex).equals("жесткая") ||
                        registry.getWorkTypes().get(workIndex).equals("штукатурный") ||
                        registry.getWorkNames().get(workIndex).equals("ФН")
                ) {
                    weeksForPrep = 1;
                }
            } catch (Exception e) {
                System.out.println(registry.getWorkNames().get(workIndex));
            }


            // Определение недель на подготовительный этап если необходимо разрешение КГИОП
            if (registry.getStatus().get(workIndex)) {
                weeksForPrep = 5;
            }

            if (
                    weeksForPrep > 0 &
                    (
                        workIndex == 0 ||
                        !registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex -1))
                    )
            ) {

                Row prepRow = sheet.createRow(sheet.getLastRowNum() + 1);

                Cell prepNumCell= prepRow.createCell(0);
                prepNumCell.setCellStyle(tableStyle);

                Cell prepNameCell= prepRow.createCell(1);
                prepNameCell.setCellStyle(nameStyle);

                int[] prep;
                if (weeksForPrep == 5) {

                    prepNameCell.setCellValue("Оформление разрешительной документации (получение ордера ГАТИ," +
                        " согласование схем ОДД, получение разрешения КГИОП и пр.)");

                    prep = new int[] {17,17,17,17,16,16};

                } else {

                    prepNameCell.setCellValue("Оформление разрешительной документации (получение ордера ГАТИ," +
                            " согласование схем ОДД и пр.)");

                    prep = new int[] {25,25,25,25};

                }

                if (isLift == 1) {
                    Cell prepLiftCell= prepRow.createCell(2);
                    prepLiftCell.setCellStyle(tableStyle);
                }

                Cell prepSumCell= prepRow.createCell(2 + isLift);
                prepSumCell.setCellStyle(costStyle);

                for (int j=0; j < prep.length; j++) {
                    Cell cell = prepRow.createCell(j + 3 + isLift);
                    cell.setCellStyle(tableStyle);
                    cell.setCellValue(prep[j]);
                }

                for (int j=prepRow.getLastCellNum(); j < sheet.getRow(6).getLastCellNum(); j++) {
                    Cell cell = prepRow.createCell(j);
                    cell.setCellStyle(tableStyle);
                }

            }

            // Строка под вид работ
            Row workRow = sheet.createRow(sheet.getLastRowNum() + 1);

            // Ячейка для номера вида работ
            Cell workNumCell = workRow.createCell(0);

            if (workIndex >0 && !registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex -1))) {
                workNum = 0;
            }

            workNum = workNum + 1;
            workNumCell.setCellValue(addressNum + "." + workNum);
            workNumCell.setCellStyle(tableStyle);

            // Ячейка для наименования вида работ
            Cell workNameCell = workRow.createCell(1);
            String workName;

            if (workStages == null) {
                switch (registry.getWorkNames().get(workIndex)) {
                    case "Лифт":
                        workName = "Ремонт или замена лифтового оборудования, ремонт лифтовых шахт";
                        break;
                    case "ТО":
                        workName = "Полное техническое освидетельствование лифта после замены лифтового оборудования";
                        break;
                    default:
                        workName = "";
                        break;
                }

                Cell workRegNumCell = workRow.createCell(2);
                workRegNumCell.setCellValue(registry.getRegNum().get(workIndex));
                workRegNumCell.setCellStyle(tableStyle);

            } else {
                workName = workStages.getWorkName();
            }

            workNameCell.setCellValue(workName);
            workNameCell.setCellStyle(workStyle);

            for (int j = workRow.getLastCellNum(); j < sheet.getRow(6).getLastCellNum(); j++) {
                Cell cell = workRow.createCell(j);
                cell.setCellStyle(tableStyle);
            }

            int designTerm = !registry.getWorkNames().get(workIndex).contains("ПД") ? registry.getDesignTerm() / 7 : 0;

            if (isLift == 1 & !registry.getWorkNames().get(workIndex).contains("ЛО_")) {

                try {
                    if (registry.getWorkNames().get(workIndex).contains("Лифт")) {

                        LiftStages.addLiftStages(
                                sheet,
                                Math.round(registry.getTerms().get(workIndex) / 7f),
                                registry.getWorkNames().get(workIndex),
                                addressNum,
                                workNum,
                                designTerm
                        );
                    } else {

                        LiftStages.addLiftStages(
                                sheet,
                                registry.getTerms().get(workIndex),
                                registry.getWorkNames().get(workIndex),
                                addressNum,
                                workNum,
                                designTerm
                        );
                    }
                } catch (Exception e) {
                    System.out.println("Ошибка в выборе лифт/ТО");
                }
            } else {
                assert workStages != null;
                Stages.addStages(
                        sheet,
                        workStages,
                        Math.round(registry.getTerms().get(workIndex) / 7f),
                        weeksForPrep,
                        addressNum,
                        workNum,
                        designTerm
                );
            }

            try {
                if (
                        workIndex < registry.getAddresses().size() - 1 &&
                                !registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex + 1)) ||
                                workIndex == registry.getAddresses().size() - 1
                ) {

                    // Строка для стоимости
                    Row workSumRow = sheet.createRow(sheet.getLastRowNum() + 1);

                    Cell workSumNumCell = workSumRow.createCell(0);
                    workSumNumCell.setCellStyle(tableStyle);

                    Cell workSumNameCell = workSumRow.createCell(1);
                    workSumNameCell.setCellValue("Итого по объекту");
                    workSumNameCell.setCellStyle(nameStyle);

                    double addressCost = 0;
                    for (int x = 0; x < registry.getCost().size(); x++) {
                        if (registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(x))) {
                            addressCost = addressCost + registry.getCost().get(x);
                        }
                    }

                    Cell workSumValCell = workSumRow.createCell(2 + isLift);
                    workSumValCell.setCellValue(addressCost);
                    workSumValCell.setCellStyle(costStyle);

                    for (int j = workSumRow.getLastCellNum(); j < sheet.getRow(6).getLastCellNum(); j++) {

                        Cell cell = workSumRow.createCell(j);
                        cell.setCellStyle(tableStyle);

                    }

                    if (
                            workIndex < registry.getAddresses().size() - 1 &&
                                    registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex + 1)) ||
                                    workIndex > 0 &&
                                            registry.getAddresses().get(workIndex).equals(registry.getAddresses().get(workIndex - 1))
                    ) {

                        // Ячейка для стоимости вида работ
                        Cell workNameSumCell = workRow.createCell(2 + isLift);
                        workNameSumCell.setCellValue(registry.getCost().get(workIndex));
                        workNameSumCell.setCellStyle(costStyle);

                    }

                } else {

                    // Ячейка для стоимости вида работ
                    Cell workNameSumCell = workRow.createCell(2 + isLift);
                    workNameSumCell.setCellValue(registry.getCost().get(workIndex));
                    workNameSumCell.setCellStyle(costStyle);

                }
            } catch (Exception e) {
                System.out.println("Ошибка в заполнении этапов");
            }
        }
    }
}
