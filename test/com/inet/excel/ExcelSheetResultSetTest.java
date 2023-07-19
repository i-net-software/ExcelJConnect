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
package com.inet.excel;

import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.File;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser;
import com.inet.excel.parser.ExcelParserTest;

public class ExcelSheetResultSetTest {

    @Test
    public void constructor_throws_exception_if_parser_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSet( null, "Sheet1", 100 ) );
    }

    @Test
    public void constructor_throws_exception_if_sheetName_is_null() {
        ExcelParser parser = getParser();
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSet( parser, null, 100 ) );
    }

    @SuppressWarnings( "resource" )
    @Test
    public void constructor_throws_exception_if_maxRowsPerBatch_is_not_greater_than_zero() {
        ExcelParser parser = getParser();
        String sheetName = "Sheet1";
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSet( parser, sheetName, -1 ) );
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSet( parser, sheetName, 0 ) );
        new ExcelSheetResultSet( parser, sheetName, 1 ); // should not throw exception
    }
    
    /** Returns parser for test purposes, which is able to read data from existing Excel document.
     * @return parser instance.
     */
    private ExcelParser getParser() {
        File resource = new File( ExcelParserTest.class.getResource( "./files/rows.xlsx" ).getPath() );
        return new ExcelParser( resource.toPath(), false );
    }
}
