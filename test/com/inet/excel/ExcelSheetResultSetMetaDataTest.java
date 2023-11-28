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

import java.util.ArrayList;

import org.junit.jupiter.api.Test;

public class ExcelSheetResultSetMetaDataTest {

    @Test
    public void constructor_should_accept_empty_list_as_columnNames_and_columnTypes() {
        new ExcelSheetResultSetMetaData( "fileName.xlsx", "Sheet1", new ArrayList<>(), new ArrayList<>() );
    }

    @Test
    public void constructor_throws_exception_if_fileName_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSetMetaData( null, "Sheet1", new ArrayList<>(), new ArrayList<>() ) );
    }

    @Test
    public void constructor_throws_exception_if_sheetName_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSetMetaData( "fileName.xlsx", null, new ArrayList<>(), new ArrayList<>() ) );
    }

    @Test
    public void constructor_throws_exception_if_columnNames_are_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSetMetaData( "fileName.xlsx", "Sheet1", null, new ArrayList<>() ) );
    }

    @Test
    public void constructor_throws_exception_if_columnTypes_are_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelSheetResultSetMetaData( "fileName.xlsx", "Sheet1", new ArrayList<>(), null ) );
    }
}
