package com.fkr.schedule.scheduleparts;

import com.fkr.schedule.beautify.Styles;
import com.fkr.schedule.domain.Registry;
import org.apache.poi.ss.usermodel.*;

public class Total {

    public static void addTotal(Sheet sheet, Registry registry, int isLift) {

        CellStyle tableStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.CENTER,
                true,
                false,
                (short) 0
        );

        CellStyle costStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.CENTER,
                true,
                false,
                (short) 4
        );

        CellStyle nameStyle = Styles.createTableStyle(
                sheet.getWorkbook(),
                HorizontalAlignment.LEFT,
                true,
                false,
                (short) 0
        );

        Row total = sheet.createRow(sheet.getLastRowNum() + 1);

        Cell totalNum = total.createCell(0);
        totalNum.setCellStyle(tableStyle);

        Cell totalName = total.createCell(1);
        totalName.setCellStyle(nameStyle);
        totalName.setCellValue("Всего по договору");

        if (isLift == 1) {
            Cell totalRegNum = total.createCell(2);
            totalRegNum.setCellStyle(tableStyle);
        }

        Double contractSum = 0d;

        for (double cost : registry.getCost()) {
            contractSum += cost;
        }

        Cell totalSum = total.createCell(2 + isLift);
        totalSum.setCellStyle(costStyle);
        totalSum.setCellValue(contractSum);

        for (int j = total.getLastCellNum(); j < sheet.getRow(6).getLastCellNum(); j++) {
            Cell cell = total.createCell(j);
            cell.setCellStyle(tableStyle);
        }

    }

}
