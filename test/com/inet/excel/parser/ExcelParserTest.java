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

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.fail;
import static java.util.Arrays.asList;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.function.Consumer;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser.ExcelParserException;

public class ExcelParserTest {

    @Test
    public void constructor_throws_exception_if_filePath_is_null() {
        boolean hasHeaderRow = true; // irrelevant for this test
        assertThrows( IllegalArgumentException.class, () -> new ExcelParser( null, hasHeaderRow ) );
    }

    @Test
    public void getFileName_returns_name_taken_from_filePath() {
        boolean hasHeaderRow = true; // irrelevant for this test
        assertEquals( "abcd.xlsx", new ExcelParser( Paths.get( "C:\\folder\\abcd.xlsx" ), hasHeaderRow ).getFileName() );
        assertEquals( "zxc.xlsx", new ExcelParser( Paths.get( "zxc.xlsx" ), hasHeaderRow ).getFileName() );
    }

    @Test
    public void getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook() {
        getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( "./files/sheet_names.xlsx", asList( "Green", "Yellow", "Red", "Blue", "Black" ) );
        getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( "./files/sheet_names_differentSheetOrder.xlsx", asList( "Yellow", "Blue", "Black", "Red", "Green" ) );
    }

    private void getSheetNames_returns_sheet_names_ordered_by_occurrence_in_workbook( String resourcePath, List<String> expectedSheetNames ) {
        File resource = new File( ExcelParserTest.class.getResource( resourcePath ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );
        assertEquals( expectedSheetNames, parser.getSheetNames() );
    }

    @Test
    public void getSheetNames_throws_exception_if_excel_file_does_not_exist() {
        boolean hasHeaderRow = true; // irrelevant for this test
        method_throws_exception_if_excel_file_does_not_exist( hasHeaderRow, ExcelParser::getSheetNames );
    }

    //TODO test: getColumnNames( null )

    @Test
    public void getColumnNames_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( true, parser -> parser.getColumnNames( "sheetName" ) );
        method_throws_exception_if_excel_file_does_not_exist( false, parser -> parser.getColumnNames( "sheetName" ) );
    }

    private void method_throws_exception_if_excel_file_does_not_exist( boolean hasHeaderRow, Consumer<ExcelParser> executable ) {
        Path path = Paths.get( "missing_file.xlsx" );
        assertFalse( Files.exists( path ) ); // precondition check
        ExcelParser parser = new ExcelParser( path, hasHeaderRow );
        try {
            executable.accept( parser );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( NoSuchFileException.class, ex.getCause().getClass() ); //TODO rethink type of exception
        }
    }

    @Test
    public void getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet() {
        getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( true );
        getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( false );
    }

    private void getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet( boolean hasHeaderRow ) {
        String nonExistingSheetName = "nonExistingSheetName";
        File resource = new File( ExcelParserTest.class.getResource( "./files/sheet_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );
        assertFalse( parser.getSheetNames().contains( nonExistingSheetName ) ); // precondition check
        try {
            parser.getColumnNames( nonExistingSheetName );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( IllegalArgumentException.class, ex.getCause().getClass() ); //TODO rethink type of exception
        }
    }

    @Test
    public void getColumnNames_returns_names_of_columns_based_on_first_row_which_is_specified_as_header() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );
        assertEquals( asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7" ), parser.getColumnNames( "Sheet1" ) );
        assertEquals( asList( "Cat", "C4", "C5", "Fun" ), parser.getColumnNames( "Sheet2" ) );
    }

    @Test
    public void getColumnNames_returns_autogenerated_names_of_columns_if_first_row_which_is_specified_as_header_is_empty() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names_firstRowIsEmpty.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );
        assertEquals( asList( "C2", "C3", "C4", "C5", "C6", "C7" ), parser.getColumnNames( "Sheet1" ) );
        assertEquals( asList( "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2" ) );
    }

    @Test
    public void getColumnNames_returns_autogenerated_names_of_columns_if_first_row_is_not_specified_as_header() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), false );
        assertEquals( asList( "C2", "C3", "C4", "C5", "C6", "C7" ), parser.getColumnNames( "Sheet1" ) );
        assertEquals( asList( "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2" ) );
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
        assertEquals( asList( "C1" ), new ExcelParser( resource.toPath(), true ).getColumnNames( "Sheet1" ) );
        assertEquals( asList( "C1" ), new ExcelParser( resource.toPath(), false ).getColumnNames( "Sheet1" ) );
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

        ExcelParser parser = new ExcelParser( resource.toPath(), true );
        assertEquals( asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7", "C8", "C9", "C10", "C11" ), parser.getColumnNames( "Sheet1" ) );
        assertEquals( asList( "C2", "Cat", "C4", "C5", "Fun" ), parser.getColumnNames( "Sheet2" ) );

        parser = new ExcelParser( resource.toPath(), false );
        assertEquals( asList( "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11" ), parser.getColumnNames( "Sheet1" ) );
        assertEquals( asList( "C2", "C3", "C4", "C5", "C6" ), parser.getColumnNames( "Sheet2" ) );
    }

    @Test
    public void getRows_returns_data_from_specified_rows() {
        List<String> emptyRow = asList( "", "", "", "", "" );

        List<String> row1 = asList( "Red", "Black", "Green", "Yellow", "Blue" );
        List<String> row2 = asList( "Cat", "", "Dog", "Bird", "" );
        List<String> row3 = asList( "", "Flower", "Garden", "Tree", "Soil" );
        List<String> row4 = emptyRow;
        List<String> row5 = asList( "Fire", "", "Ice", "", "Water" );
        List<String> row6 = asList( "Sun", "", "", "", "" );
        List<String> row7 = asList( "", "", "", "", "Moon" );
        List<String> row8 = asList( "One", "Two", "Three", "Four", "Five" );

        String sheetName = "Sheet1";
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );

        assertEquals( asList( row1, row2, row3, row4, row5, row6, row7, row8 ), parser.getRows( sheetName, 1, 8 ) );

        assertEquals( asList( row1 ), parser.getRows( sheetName, 1, 1 ) );
        assertEquals( asList( row3 ), parser.getRows( sheetName, 3, 3 ) );
        assertEquals( asList( row4, row5, row6 ), parser.getRows( sheetName, 4, 6 ) );

        assertEquals( asList( row8, emptyRow ), parser.getRows( sheetName, 8, 9 ) );
        assertEquals( asList( emptyRow, emptyRow, emptyRow ), parser.getRows( sheetName, 55, 57 ) );
    }

    @Test
    public void getRows_throws_exception_if_specified_column_indexes_are_invalid() {
        assertGetRowsThrowsException( 0, 1 );
        assertGetRowsThrowsException( 1, 0 );
        assertGetRowsThrowsException( 2, 1 );
    }

    private void assertGetRowsThrowsException( int firstRowIndex, int lastRowIndex ) {
        String sheetName = "Sheet1";
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );

        assertDoesNotThrow( () -> parser.getRows( sheetName, 1, 1 ) ); // precondition check
        assertThrows( IllegalArgumentException.class, () -> parser.getRows( sheetName, firstRowIndex, lastRowIndex ) );
    }

    //TODO test: getRows( null, 1, 2 )

    @Test
    public void getRows_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( true, parser -> parser.getRows( "sheetName", 1, 2 ) );
        method_throws_exception_if_excel_file_does_not_exist( false, parser -> parser.getRows( "sheetName", 1, 2 ) );
    }

    @Test
    public void getRows_throws_exception_if_workbook_does_not_include_specified_sheet() {
        String nonExistingSheetName = "nonExistingSheetName";
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );
        assertFalse( parser.getSheetNames().contains( nonExistingSheetName ) ); // precondition check
        try {
            parser.getRows( nonExistingSheetName, 1, 2 );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( IllegalArgumentException.class, ex.getCause().getClass() ); //TODO rethink type of exception
        }
    }

    @Test
    public void getRowCount_returns_number_of_rows_included_in_specified_sheet() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/row_count.xlsx" ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );

        assertEquals( 3, parser.getRowCount( "Sheet1" ) );
        assertEquals( 17, parser.getRowCount( "Sheet2" ) );
        assertEquals( 0, parser.getRowCount( "Sheet3" ) );
    }

    //TODO test: getRowCount( null )

    @Test
    public void getRowCount_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( true, parser -> parser.getRowCount( "sheetName" ) );
        method_throws_exception_if_excel_file_does_not_exist( false, parser -> parser.getRowCount( "sheetName" ) );
    }
}
