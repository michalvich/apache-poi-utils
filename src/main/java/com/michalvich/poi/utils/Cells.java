package com.michalvich.poi.utils;

import org.apache.poi.ss.usermodel.Cell;

import java.text.DecimalFormat;

public class Cells {

    public static String getValueAsString(Cell cell) {

        switch (cell.getCellType()) {

            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case Cell.CELL_TYPE_NUMERIC:
                double value = cell.getNumericCellValue();
                return new DecimalFormat("#").format(value);
            case Cell.CELL_TYPE_BLANK:
                return "";

        }

        throw new IllegalArgumentException(String.format("Unexpected cell type %s", cell.getCellType()));

    }

}
