package com.michalvich.poi.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.assertEquals;

public class CellsTest {

    private Cell cell;

    private Workbook workbook;

    @Before
    public void setUp() throws Exception {
        workbook = createWorkbook();
        cell = workbook.getSheetAt(0).getRow(0).getCell(0);
    }

    @Test
    public void should_return_string_for_cell_type_string() throws Exception {

        final String expectedValue = "string";

        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellValue(expectedValue);

        assertEquals(expectedValue, Cells.getValueAsString(cell));

    }

    @Test
    public void should_return_string_for_cell_type_formula() throws Exception {

        final String expectedValue = "12";

        cell.setCellType(Cell.CELL_TYPE_FORMULA);
        cell.setCellFormula("12");
        cell.setCellValue(expectedValue);

        assertEquals(expectedValue.toString(), Cells.getValueAsString(cell));

    }


    @Test
    public void should_return_string_for_cell_type_numeric() throws Exception {

        final Double expectedValue = 12.0;

        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
        cell.setCellValue(expectedValue);

        assertEquals("12", Cells.getValueAsString(cell));

    }

    private Workbook createWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        workbook.createSheet().createRow(0).createCell(0);
        return workbook;
    }

}
