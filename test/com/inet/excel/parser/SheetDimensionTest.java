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
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Arrays;
import java.util.List;

import org.junit.Test;

public class SheetDimensionTest {

    @Test
    public void constructor_doethrows_exception_if_specified_column_indexes_are_invalid() {
        assertConstructorCreatesInstance( 1, 1 );
        assertConstructorCreatesInstance( 1, 2 );
        assertConstructorCreatesInstance( 2, 2 );
        assertConstructorCreatesInstance( 5, 88 );
    }

    private void assertConstructorCreatesInstance( int firstColumnIndex, int lastColumnIndex ) {
        SheetDimension dimension = new SheetDimension( firstColumnIndex, lastColumnIndex );
        assertEquals( firstColumnIndex, dimension.getFirstColumnIndex() );
        assertEquals( lastColumnIndex, dimension.getLastColumnIndex() );
    }

    @Test
    public void constructor_throws_exception_if_specified_column_indexes_are_invalid() {
        assertConstructorThrowsException( 0, 1 );
        assertConstructorThrowsException( 1, 0 );
        assertConstructorThrowsException( 2, 1 );
    }

    private void assertConstructorThrowsException( int firstColumnIndex, int lastColumnIndex ) {
        assertThrows( IllegalArgumentException.class, () -> new SheetDimension( firstColumnIndex, lastColumnIndex ) );
    }

    @Test
    public void parse_creates_instance_based_on_dimension_ref() {
        assertDimension( 2, 2, "B2" );
        assertDimension( 5, 5, "E1" );
        assertDimension( 1, 3, "A3:C6" );
        assertDimension( 4, 26, "D7:Z15" );
        assertDimension( 27, 29, "AA3:AC3" );
    }

    private void assertDimension( int expectedFirst, int expectedLast, String toParse ) {
        SheetDimension dimension = SheetDimension.parse( toParse );
        assertEquals( expectedFirst, dimension.getFirstColumnIndex() );
        assertEquals( expectedLast, dimension.getLastColumnIndex() );
    }

    @Test
    public void parse_returns_null_if_dimension_ref_is_invalid() {
        List<String> invalidRefs = Arrays.asList( null, //
                                                  "", //
                                                  "   ", //
                                                  "+", //
                                                  "()", //
                                                  "3", //
                                                  ":", //
                                                  "A3:44", //
                                                  "A3:", //
                                                  "5:F4", //
                                                  ":D3", //
                                                  "22:33", //
                                                  "B2:C5:D8" );

        for( String ref : invalidRefs ) {
            assertNull( SheetDimension.parse( ref ), "ref=\"" + ref + "\"" );
        }
    }

    @Test
    public void getColumnIndexFromCellRef_returns_column_index_from_valid_cell_ref() {
        assertEquals( 1, SheetDimension.getColumnIndexFromCellRef( "A3" ) );
        assertEquals( 18, SheetDimension.getColumnIndexFromCellRef( "R7" ) );
        assertEquals( 26, SheetDimension.getColumnIndexFromCellRef( "Z1" ) );
        assertEquals( 27, SheetDimension.getColumnIndexFromCellRef( "AA1" ) );
        assertEquals( 110, SheetDimension.getColumnIndexFromCellRef( "DF55" ) );
    }

    @Test
    public void getColumnIndexFromCellRef_returns_zero_if_cell_ref_is_invalid() {
        List<String> invalidRefs = Arrays.asList( null, //
                                                  "", //
                                                  "   ", //
                                                  "+", //
                                                  "()", //
                                                  "3" );

        for( String ref : invalidRefs ) {
            assertEquals( 0, SheetDimension.getColumnIndexFromCellRef( ref ), "ref=\"" + ref + "\"" );
        }
    }
}
