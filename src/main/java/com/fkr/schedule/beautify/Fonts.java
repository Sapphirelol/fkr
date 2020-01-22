package com.fkr.schedule.beautify;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class Fonts {

    public static Font createTitleBoldFont(Workbook wb) {

        Font font = wb.createFont();
        font.setFontHeightInPoints((short)12);
        font.setFontName("Times New Roman");
        font.setBold(true);

        return font;
    }

    public static Font createTitleFont(Workbook wb) {

        Font font = wb.createFont();
        font.setFontHeightInPoints((short)12);
        font.setFontName("Times New Roman");

        return font;
    }

    public static Font createTableFont(Workbook wb, boolean isBold, boolean isItalic) {

        Font font = wb.createFont();
        font.setFontHeightInPoints((short)10);
        font.setFontName("Times New Roman");
        font.setBold(isBold);
        font.setItalic(isItalic);

        return font;
    }

}
