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
import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;

public class ExcelDatabaseResultSetTest {

    @SuppressWarnings( "resource" )
    @Test
    public void constructor_should_accept_empty_lists_as_columnNames_and_rows() {
        new ExcelDatabaseResultSet( new ArrayList<>(), new ArrayList<>() );
    }

    @Test
    public void constructor_throws_exception_if_columnNames_are_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelDatabaseResultSet( null, new ArrayList<>() ) );
    }

    @Test
    public void constructor_throws_exception_if_rows_are_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelDatabaseResultSet( new ArrayList<>(), null ) );
    }

    @Test
    public void constructor_throws_exception_if_number_of_values_in_all_rows_does_not_match_number_of_columns() {
        List<String> columnNames = Arrays.asList( "first", "second", "third" );

        List<List<Object>> rowsWithNotEnoughData = Arrays.asList( Arrays.asList( "A", "S", "D" ), //
                                                                  Arrays.asList( "1", "2" ), //
                                                                  Arrays.asList( "Z", "X", "C" ) );
        assertThrows( IllegalArgumentException.class, () -> new ExcelDatabaseResultSet( columnNames, rowsWithNotEnoughData ) );

        List<List<Object>> rowsWithTooMuchData = Arrays.asList( Arrays.asList( "A", "S", "D" ), //
                                                                Arrays.asList( "1", "2", "3" ), //
                                                                Arrays.asList( "Z", "X", "C", "V" ) );
        assertThrows( IllegalArgumentException.class, () -> new ExcelDatabaseResultSet( columnNames, rowsWithTooMuchData ) );
    }
}
