package com.lagerweij.cukesheets;

import bad.robot.excel.DoubleCell;
import bad.robot.excel.StringCell;
import bad.robot.excel.valuetypes.Coordinate;
import bad.robot.excel.valuetypes.ExcelColumnIndex;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;

import static bad.robot.excel.matchers.CellType.*;
import static bad.robot.excel.matchers.WorkbookMatcher.*;
import static org.hamcrest.MatcherAssert.*;
import static org.hamcrest.Matchers.equalTo;
import static org.hamcrest.Matchers.not;

public class CukesSheetStepdefs {

    Workbook inputSheet;
    Workbook outputSheet;

    @Given("^the input sheet input_sheet_(\\d+).xls$")
    public void the_input_sheet_input_sheet_xls(int arg1) throws Throwable {
        InputStream file = this.getClass().getResourceAsStream("/input_sheet_1.xls");
        inputSheet = new HSSFWorkbook(file);
    }

    @When("^the file is processed$")
    public void the_file_is_processed() throws Throwable {
        // Here's where output files are generated
    }

    @Then("^the output is exactly the same as (.+)$")
    public void the_output_is_exactly_the_same_as_baseline_sheet__matching_xls(String filename) throws Throwable {
        InputStream file = this.getClass().getResourceAsStream("/" + filename);
        Workbook baseline = new HSSFWorkbook(file);
        assertThat(inputSheet, sameWorkbook(baseline));
    }

    @Then("^the output is not the same as baseline_sheet_1_not_matching.xls$")
    public void the_output_is_not_the_same_as_baseline_sheet__not_matching_xls() throws Throwable {
        InputStream file = this.getClass().getResourceAsStream("/baseline_sheet_1_not_matching.xls");
        Workbook baseline = new HSSFWorkbook(file);
        assertThat(inputSheet, not(sameWorkbook(baseline)));
    }

    @Then("^the output is generated in (.+)$")
    public void the_output_is_generated_in_baseline_sheet__matching_xls(String sheetName) throws Throwable {
        InputStream file = this.getClass().getResourceAsStream("/baseline_sheet_1_not_matching.xls");
        outputSheet = new HSSFWorkbook(file);
    }

    @Then("^cell at coordinate (\\d+), (\\d+) has numeric value ([\\d\\.]+)$")
    public void cell_at_coordinate_C_has_value_(int columnIndex, int rowIndex, double value) throws Throwable {
        Coordinate c = Coordinate.coordinate(ExcelColumnIndex.from(columnIndex), rowIndex);
        Sheet sheet = outputSheet.getSheetAt(c.getSheet().value());
        Row row = sheet.getRow(c.getRow().value());
        Cell poiCell = row.getCell(c.getColumn().value());
        DoubleCell cell = (DoubleCell) adaptPoi(poiCell);
        assertThat(cell, equalTo(new DoubleCell(value)));
    }

    @Then("^cell at coordinate (\\d+), (\\d+) has string value (.+)$")
    public void cell_at_coordinate_C_has_value_(int columnIndex, int rowIndex, String value) throws Throwable {
        Coordinate c = Coordinate.coordinate(ExcelColumnIndex.from(columnIndex), rowIndex);
        Sheet sheet = outputSheet.getSheetAt(c.getSheet().value());
        Row row = sheet.getRow(c.getRow().value());
        Cell poiCell = row.getCell(c.getColumn().value());
        StringCell cell = (StringCell) adaptPoi(poiCell);
        assertThat(cell, equalTo(new StringCell(value)));
    }
}
