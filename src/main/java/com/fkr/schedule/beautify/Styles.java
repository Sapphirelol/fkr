package com.fkr.schedule.beautify;

import org.apache.poi.ss.usermodel.*;

public class Styles {

    public static CellStyle createTitleBoldStyle(Workbook wb) {

        CellStyle style = wb.createCellStyle();

        style.setFont(Fonts.createTitleBoldFont(wb));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return style;
    }

    public static CellStyle createTitleStyle(Workbook wb) {

        CellStyle style = wb.createCellStyle();

        style.setFont(Fonts.createTitleFont(wb));
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return style;
    }

    public static CellStyle createTableStyle(Workbook wb,
                                             HorizontalAlignment alignment,
                                             boolean isBold,
                                             boolean isItalic,
                                             short dataFormat
    ) {

        CellStyle style = wb.createCellStyle();
        DataFormat format = wb.createDataFormat();

        style.setFont(Fonts.createTableFont(wb, isBold, isItalic));
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        if (dataFormat != 0) {
            style.setDataFormat(dataFormat);
        }

        return style;
    }

}
