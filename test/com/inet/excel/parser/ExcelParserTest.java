/*
 * Copyright 2023 i-net software
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.inet.excel.parser;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser.ExcelParserException;

public class ExcelParserTest {

    @Test
    public void constructor_throws_exception_if_filePath_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelParser( null ) );
    }

    @Test
    public void getFileName_returns_name_taken_from_filePath() {
        assertEquals( "abcd.xlsx", new ExcelParser( Paths.get( "C:\\folder\\abcd.xlsx" ) ).getFileName() );
        assertEquals( "zxc.xlsx", new ExcelParser( Paths.get( "zxc.xlsx" ) ).getFileName() );
    }

    @Test
    public void getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook() {
        getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( "./files/sheet_names.xlsx", Arrays.asList( "Green", "Yellow", "Red", "Blue", "Black" ) );
        getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( "./files/sheet_names_differentSheetOrder.xlsx", Arrays.asList( "Yellow", "Blue", "Black", "Red", "Green" ) );
    }

    private void getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( String resourcePath, List<String> expectedSheetNames ) {
        File resource = new File( ExcelParserTest.class.getResource( resourcePath ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertEquals( expectedSheetNames, parser.getSheetNames() );
    }

    @Test
    public void getSheetNames_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( ExcelParser::getSheetNames );
    }

    @Test
    public void getColumnNames_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getColumnNames( "sheetName", true ) );
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getColumnNames( "sheetName", false ) );
    }

    @Test
    public void getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet() {
        getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( true );
        getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( false );
    }

    private void getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( boolean hasHeaderRow ) {
        String nonExistingSheetName = "nonExistingSheetName";
        File resource = new File( ExcelParserTest.class.getResource( "./files/sheet_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertFalse( parser.getSheetNames().contains( nonExistingSheetName ) ); // precondition check
        try {
            parser.getColumnNames( nonExistingSheetName, hasHeaderRow );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( IllegalArgumentException.class, ex.getCause().getClass() ); //TODO rethink type of exception
        }
    }

    @Test
    public void getColumnNames_returns_names_of_columns_based_on_first_row_which_is_specified_as_header() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertEquals( Arrays.asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7" ), parser.getColumnNames( "Sheet1", true ) );
        assertEquals( Arrays.asList( "Cat", "C4", "C5", "Fun" ), parser.getColumnNames( "Sheet2", true ) );
    }

    @Test
    public void getColumnNames_returns_autogenerated_names_of_columns_if_first_row_which_is_specified_as_header_is_empty() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names_firstRowIsEmpty.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertEquals( Arrays.asList( "C2", "C3", "C4", "C5", "C6", "C7" ), parser.getColumnNames( "Sheet1", true ) );
        assertEquals( Arrays.asList( "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2", true ) );
    }

    @Test
    public void getColumnNames_returns_autogenerated_names_of_columns_if_first_row_is_not_specified_as_header() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertEquals( Arrays.asList( "C2", "C3", "C4", "C5", "C6", "C7" ), parser.getColumnNames( "Sheet1", false ) );
        assertEquals( Arrays.asList( "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2", false ) );
    }

    @Test
    public void getColumnNames_returns_autogenerated_name_of_single_column_when_sheet_is_empty() {
        getColumnNames_returns_autogenerated_name_of_single_column_when_sheet_is_empty( "./files/empty_sheet.xlsx" );
    }

    @Test
    public void getColumnNames_returns_autogenerated_name_of_single_column_when_sheet_is_empty_even_if_sheet_dimensions_are_missing() {
        getColumnNames_returns_autogenerated_name_of_single_column_when_sheet_is_empty( "./files/empty_sheet_missingSheetDimensions.xlsx" );
    }

    private void getColumnNames_returns_autogenerated_name_of_single_column_when_sheet_is_empty( String resourcePath ) {
        File resource = new File( ExcelParserTest.class.getResource( resourcePath ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );
        assertEquals( Arrays.asList( "C1" ), parser.getColumnNames( "Sheet1", true ) );
        assertEquals( Arrays.asList( "C1" ), parser.getColumnNames( "Sheet1", false ) );
    }

    @Test
    public void getColumnNames_returns_names_of_columns_even_if_sheet_dimensions_are_missing() {
        getColumnNames_returns_names_of_columns_even_if_optimization_data_are_missing( "./files/column_names_missingSheetDimensions_differentRowSpans.xlsx" );
    }

    @Test
    public void getColumnNames_returns_names_of_columns_even_if_sheet_dimensions_are_missing_and_row_spans_are_incomplete() {
        getColumnNames_returns_names_of_columns_even_if_optimization_data_are_missing( "./files/column_names_missingSheetDimensions_incompleteRowSpans.xlsx" );
    }

    @Test
    public void getColumnNames_returns_names_of_columns_even_if_sheet_dimensions_and_row_spans_are_missing() {
        getColumnNames_returns_names_of_columns_even_if_optimization_data_are_missing( "./files/column_names_missingSheetDimensions_missingRowSpans.xlsx" );
    }

    private void getColumnNames_returns_names_of_columns_even_if_optimization_data_are_missing( String resourcePath ) {
        File resource = new File( ExcelParserTest.class.getResource( resourcePath ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath() );

        assertEquals( Arrays.asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7", "C8", "C9", "C10", "C11" ), parser.getColumnNames( "Sheet1", true ) );
        assertEquals( Arrays.asList( "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11" ), parser.getColumnNames( "Sheet1", false ) );

        assertEquals( Arrays.asList( "C2", "Cat", "C4", "C5", "Fun" ), parser.getColumnNames( "Sheet2", true ) );
        assertEquals( Arrays.asList( "C2", "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2", false ) );
    }

    private void method_throws_exception_if_excel_file_does_not_exist( Consumer<ExcelParser> executable ) {
        Path path = Paths.get( "missing_file.xlsx" );
        assertFalse( Files.exists( path ) ); // precondition check
        ExcelParser parser = new ExcelParser( path );
        try {
            executable.accept( parser );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( NoSuchFileException.class, ex.getCause().getClass() ); //TODO rethink type of exception
        }
    }
}
