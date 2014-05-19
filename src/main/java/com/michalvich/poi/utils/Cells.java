package com.michalvich.poi.utils;

import org.apache.poi.ss.usermodel.Cell;

public class Cells {

    public static String getValueAsString(Cell cell) {

        switch (cell.getCellType()) {

            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return Double.valueOf(cell.getNumericCellValue()).toString();
            case Cell.CELL_TYPE_BLANK:
                return "";

        }

        throw new IllegalArgumentException(String.format("Unexpected cell type %s", cell.getCellType()));

    }

}
