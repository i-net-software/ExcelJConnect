/*
 * Copyright 2023 - 2024 i-net software
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
package com.inet.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.File;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser;
import com.inet.excel.parser.ExcelParserTest;

public class ExcelDatabaseMetaDataTest {

    @Test
    public void constructor_throws_exception_if_parser_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelDatabaseMetaData( null ) );
    }

    @Test
    public void getProcedures_returns_information_about_procedures_representing_all_sheets_from_excel_document() throws SQLException {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );
        ExcelDatabaseMetaData metaData = new ExcelDatabaseMetaData( parser );

        List<String> procedureNames = new ArrayList<>();
        ResultSet rs = metaData.getProcedures( null, null, null );
        while( rs.next() ) {
            procedureNames.add( rs.getString( "PROCEDURE_NAME" ) );
        }

        List<String> expectedProcedureNames = Arrays.asList( "Sheet1", "Sheet2" );
        assertEquals( expectedProcedureNames, procedureNames );
    }

    @Test
    public void getProcedureColumns_returns_information_about_columns_from_all_sheets() throws SQLException {
        Map<String, List<String>> expectedProcedureColumns = new HashMap<>();
        expectedProcedureColumns.put( "Sheet1", Arrays.asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7" ) );
        expectedProcedureColumns.put( "Sheet2", Arrays.asList( "Cat", "C4", "C5", "Fun" ) );
        getProcedureColumns_returns_information_about_columns( null, expectedProcedureColumns );
    }

    @Test
    public void getProcedureColumns_returns_information_about_columns_from_specified_sheet() throws SQLException {
        getProcedureColumns_returns_information_about_columns( "Sheet1", Collections.singletonMap( "Sheet1", Arrays.asList( "C2", "Chocolate", "C4", "Egg", "Forest", "C7" ) ) );
        getProcedureColumns_returns_information_about_columns( "Sheet2", Collections.singletonMap( "Sheet2", Arrays.asList( "Cat", "C4", "C5", "Fun" ) ) );
    }

    private void getProcedureColumns_returns_information_about_columns( String procedureNamePattern, Map<String, List<String>> expectedProcedureColumns ) throws SQLException {
        File resource = new File( ExcelParserTest.class.getResource( "./files/column_names.xlsx" ).getPath() );
        ExcelParser parser = new ExcelParser( resource.toPath(), true );
        ExcelDatabaseMetaData metaData = new ExcelDatabaseMetaData( parser );

        Map<String, List<String>> procedureColumns = new HashMap<>();
        ResultSet rs = metaData.getProcedureColumns( null, null, procedureNamePattern, null );
        while( rs.next() ) {
            String procedureName = rs.getString( "PROCEDURE_NAME" );
            String columnName = rs.getString( "COLUMN_NAME" );
            procedureColumns.computeIfAbsent( procedureName, k -> new ArrayList<>() ).add( columnName );
        }

        assertEquals( expectedProcedureColumns, procedureColumns );
    }
}
