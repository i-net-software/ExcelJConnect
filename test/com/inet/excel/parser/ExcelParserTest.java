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
import java.sql.Date;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.function.BiConsumer;
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
        assertEquals( "abcd.xlsx", new ExcelParser( Paths.get( "C:" + File.separator + "folder" + File.separator + "abcd.xlsx" ), hasHeaderRow ).getFileName() );
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
        method_throws_exception_if_excel_file_does_not_exist( ExcelParser::getSheetNames );
    }

    @Test
    public void getColumnNames_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getColumnNames( "sheetName" ) );
    }

    private void method_throws_exception_if_excel_file_does_not_exist( Consumer<ExcelParser> executable ) {
        Path path = Paths.get( "missing_file.xlsx" );
        assertFalse( Files.exists( path ) ); // precondition check
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( path, hasHeaderRow );
        try {
            executable.accept( parser );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( NoSuchFileException.class, ex.getCause().getClass() );
        }
    }

    @Test
    public void getColumnNames_throws_exception_if_sheet_is_null() {
        method_throws_exception_if_sheet_is_invalid( null, (parser,sheetName) -> parser.getColumnNames( sheetName ) );
    }

    private void method_throws_exception_if_sheet_is_invalid( String sheetName, BiConsumer<ExcelParser,String> executable ) {
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        boolean hasHeaderRow = true; // irrelevant for this test
        ExcelParser parser = new ExcelParser( resource.toPath(), hasHeaderRow );
        assertFalse( parser.getSheetNames().contains( sheetName ) ); // precondition check
        try {
            executable.accept( parser, sheetName );
            fail( "expected exception" );
        } catch( ExcelParserException ex ) {
            assertEquals( IllegalArgumentException.class, ex.getCause().getClass() );
        }
    }

    @Test
    public void getColumnNames_throws_exception_if_workbook_does_not_include_specified_sheet() {
        method_throws_exception_if_sheet_is_invalid( "nonExistingSheetName", (parser,sheetName) -> parser.getColumnNames( sheetName ) );
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
    public void getRows_returns_data_from_specified_rows_of_document_without_header_row() {
        List<String> emptyRow = asList( null, null, null, null, null );

        List<String> row1 = asList( "Red", "Black", "Green", "Yellow", "Blue" );
        List<String> row2 = asList( "Cat", null, "Dog", "Bird", null );
        List<String> row3 = asList( null, "Flower", "Garden", "Tree", "Soil" );
        List<String> row4 = emptyRow;
        List<String> row5 = asList( "Fire", null, "Ice", null, "Water" );
        List<String> row6 = asList( "Sun", null, null, null, null );
        List<String> row7 = asList( null, null, null, null, "Moon" );
        List<String> row8 = asList( "One", "Two", "Three", "Four", "Five" );

        String sheetName = "Sheet1";
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), false );

        assertEquals( asList( row1, row2, row3, row4, row5, row6, row7, row8 ), parser.getRows( sheetName, 1, 8 ) );

        assertEquals( asList( row1 ), parser.getRows( sheetName, 1, 1 ) );
        assertEquals( asList( row3 ), parser.getRows( sheetName, 3, 3 ) );
        assertEquals( asList( row4, row5, row6 ), parser.getRows( sheetName, 4, 6 ) );

        assertEquals( asList( row8, emptyRow ), parser.getRows( sheetName, 8, 9 ) );
        assertEquals( asList( emptyRow, emptyRow, emptyRow ), parser.getRows( sheetName, 55, 57 ) );
    }

    @Test
    public void getRows_returns_data_from_specified_rows_of_document_with_header_row() {
        List<String> emptyRow = asList( null, null, null, null, null );

        List<String> row1 = asList( "Cat", null, "Dog", "Bird", null );
        List<String> row2 = asList( null, "Flower", "Garden", "Tree", "Soil" );
        List<String> row3 = emptyRow;
        List<String> row4 = asList( "Fire", null, "Ice", null, "Water" );
        List<String> row5 = asList( "Sun", null, null, null, null );
        List<String> row6 = asList( null, null, null, null, "Moon" );
        List<String> row7 = asList( "One", "Two", "Three", "Four", "Five" );

        String sheetName = "Sheet1";
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );

        assertEquals( asList( row1, row2, row3, row4, row5, row6, row7 ), parser.getRows( sheetName, 1, 7 ) );

        assertEquals( asList( row1 ), parser.getRows( sheetName, 1, 1 ) );
        assertEquals( asList( row3 ), parser.getRows( sheetName, 3, 3 ) );
        assertEquals( asList( row4, row5, row6 ), parser.getRows( sheetName, 4, 6 ) );

        assertEquals( asList( row7, emptyRow ), parser.getRows( sheetName, 7, 8 ) );
        assertEquals( asList( emptyRow, emptyRow, emptyRow ), parser.getRows( sheetName, 55, 57 ) );
    }

    @Test
    public void getRows_returns_data_of_various_types() throws ParseException {
        File resource = new File( ExcelParserTest.class.getResource( "./files/various_data_types.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), false );

        List<List<Object>> expectedRows = asList( asList( "125" ), // $125.0000 
                                                  asList( "222" ), // 222
                                                  asList( "333" ), // 333.00
                                                  asList( "444" ), // $444.00
                                                  asList( "555" ), // $555.00 
                                                  asList( "1.1000000000000001" ), // 110.00%
                                                  asList( "2.4" ), // 2 2/5
                                                  asList( "10" ), // 1.00E+01
                                                  asList( "#DIV/0!" ), // #DIV/0!
                                                  asList( "555" ) ); // 555.00

        assertEquals( expectedRows, parser.getRows( "Sheet1", 1, 10 ) );
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

    @Test
    public void getRows_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getRows( "sheetName", 1, 2 ) );
    }

    @Test
    public void getRows_throws_exception_if_sheet_is_null() {
        method_throws_exception_if_sheet_is_invalid( null, (parser,sheetName) -> parser.getRows( sheetName, 1, 2 ) );
    }

    @Test
    public void getRows_throws_exception_if_workbook_does_not_include_specified_sheet() {
        method_throws_exception_if_sheet_is_invalid( "nonExistingSheetName", (parser,sheetName) -> parser.getRows( sheetName, 1, 2 ) );
    }

    @Test
    public void getRowCount_returns_number_of_rows_included_in_specified_sheet() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/row_count.xlsx" ).getPath() );

        ExcelParser parser = new ExcelParser( resource.toPath(), false );
        assertEquals( 3, parser.getRowCount( "Sheet1" ) );
        assertEquals( 17, parser.getRowCount( "Sheet2" ) );
        assertEquals( 0, parser.getRowCount( "Sheet3" ) );

        parser = new ExcelParser( resource.toPath(), true );
        assertEquals( 2, parser.getRowCount( "Sheet1" ) );
        assertEquals( 16, parser.getRowCount( "Sheet2" ) );
        assertEquals( 0, parser.getRowCount( "Sheet3" ) );
    }

    @Test
    public void getRowCount_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getRowCount( "sheetName" ) );
    }

    @Test
    public void getRowCount_throws_exception_if_sheet_is_null() {
        method_throws_exception_if_sheet_is_invalid( null, (parser,sheetName) -> parser.getRowCount( sheetName ) );
    }

    @Test
    public void getRowCount_throws_exception_if_workbook_does_not_include_specified_sheet() {
        method_throws_exception_if_sheet_is_invalid( "nonExistingSheetName", (parser,sheetName) -> parser.getRowCount( sheetName ) );
    }

    @Test
    public void values_representing_date_or_time_are_provided_as_instances_of_time_or_date_or_timestamp() throws ParseException {
        String sheetName = "Sheet1";
        File resource = new File( ExcelParserTest.class.getResource( "./files/dates.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), false );

        List<Object> row1 = asList( new Timestamp( new SimpleDateFormat( "MM/dd/yyyy" ).parse( "3/17/1915" ).getTime() ), null, null, null );
        List<Object> row2 = asList( new Date( new SimpleDateFormat( "MM/dd/yyyy" ).parse( "4/16/1921" ).getTime() ), null, null, null );
        List<Object> row3 = asList( new Timestamp( new SimpleDateFormat( "MM/dd/yyyy hh:mm:ss a" ).parse( "5/17/1927 12:00:00 PM" ).getTime() ), null, null, null );
        List<Object> row4 = asList( "5/17/1927 12:00:00 PM", null, null, null ); // cell contains text
        assertEquals( asList( row1, row2, row3, row4 ), parser.getRows( sheetName, 1, 4 ) );

        long millis = new SimpleDateFormat( "MM/dd/yyyy" ).parse( "10/1/2024" ).getTime();
        Date date = new Date( millis );
        Timestamp timestamp = new Timestamp( millis );
        Time time = new Time( millis );

        for( int rowIndex = 5; rowIndex <= 8; rowIndex++ ) {
            String msg = "row " + rowIndex;
            assertEquals( asList( asList( date, date, date, date ) ), parser.getRows( sheetName, rowIndex, rowIndex ), msg );
        }
        assertEquals( asList( asList( date, timestamp, timestamp, timestamp ) ), parser.getRows( sheetName, 9, 9 ) );
        assertEquals( asList( asList( time, time, time, time ) ), parser.getRows( sheetName, 10, 10 ) );
    }

    @Test
    public void getColumnTypes_throws_exception_if_excel_file_does_not_exist() {
        method_throws_exception_if_excel_file_does_not_exist( parser -> parser.getColumnTypes( "sheetName" ) );
    }

    @Test
    public void getColumnTypes_throws_exception_if_sheet_is_null() {
        method_throws_exception_if_sheet_is_invalid( null, (parser,sheetName) -> parser.getColumnTypes( sheetName ) );
    }

    @Test
    public void getColumnTypes_throws_exception_if_workbook_does_not_include_specified_sheet() {
        method_throws_exception_if_sheet_is_invalid( "nonExistingSheetName", (parser,sheetName) -> parser.getColumnTypes( sheetName ) );
    }

    @Test
    public void getColumnTypes_returns_value_types_for_columns_from_specified_sheet() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_types.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );

        assertEquals( asList( ValueType.NUMBER, ValueType.DATE, ValueType.TIME, ValueType.TIMESTAMP, ValueType.NUMBER, ValueType.VARCHAR ), parser.getColumnTypes( "SingleType" ) );

        assertEquals( asList( ValueType.VARCHAR, ValueType.VARCHAR, ValueType.VARCHAR, ValueType.VARCHAR, ValueType.NUMBER, ValueType.NUMBER, ValueType.VARCHAR, ValueType.NUMBER ), parser.getColumnTypes( "MixedTypes" ) );

        assertEquals( asList( ValueType.VARCHAR, ValueType.DATE ), parser.getColumnTypes( "CellLimit" ) );

        assertEquals( asList( ValueType.VARCHAR, ValueType.VARCHAR, ValueType.DATE ), parser.getColumnTypes( "RowLimit" ) );
    }
}
