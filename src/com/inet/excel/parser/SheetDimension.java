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
package com.inet.excel.parser;

/** Immutable container for information about dimension of Excel sheet.
 */
public class SheetDimension {

    private final int firstColumnIndex;
    private final int lastColumnIndex;

    /** Creates new immutable container for information about dimension of Excel sheet.
     * @param firstColumnIndex index of first column in sheet (inclusive). Minimum value is 1.
     * @param lastColumnIndex index of last column in sheet (inclusive). Minimum value is 1.
     * @throws IllegalArgumentException if one of specified indexes is smaller than 1; if first index is greater than last index.
     */
    public SheetDimension( int firstColumnIndex, int lastColumnIndex ) {
        if( firstColumnIndex < 1 ) {
            throw new IllegalArgumentException( "firstColumnIndex must be greater than zero" );
        }
        if( lastColumnIndex < 1 ) {
            throw new IllegalArgumentException( "lastColumnIndex must be greater than zero" );
        }
        if( firstColumnIndex > lastColumnIndex ) {
            throw new IllegalArgumentException( "firstColumnIndex  must be smaller than or equal to lastColumnIndex" );
        }
        this.firstColumnIndex = firstColumnIndex;
        this.lastColumnIndex = lastColumnIndex;
    }

    /** Returns instance representing information about dimension of excel sheet retrieved from provided reference.
     * @param ref reference to retrieve sheet dimension from. 
     * @return instance representing information about dimension of excel sheet or null, in case of invalid reference.
     */
    public static SheetDimension parse( String ref ) {
        if( ref == null ) {
            return null;
        }
        try {
            String[] range = ref.split( ":" );
            if( range.length == 1 && !ref.contains( ":" ) ) {
                int index = getColumnIndexFromCellRef( range[0] );
                if( index == 0 ) {
                    return null;
                }
                return new SheetDimension( index, index );
            } else if( range.length == 2 ) {
                int first = getColumnIndexFromCellRef( range[0] );
                int last = getColumnIndexFromCellRef( range[1] );
                if( first == 0 || last == 0 ) {
                    return null;
                }
                return new SheetDimension( first, last );
            } else {
                return null; // ignores data if invalid
            }
        } catch( Exception ex ) {
            return null; // ignores data if it can not be parsed 
        }
    }

    /** Returns index of column from specified cell reference. In case of invalid cell reference, it will return zero.
     * Examples of valid cell references: "A1", "A7", "G14", "BB44", "ASD1".
     * Examples of invalid cell references: null, "", "   ", "5", "$".
     * @param cellRef cell reference to retrieve cell index from.
     * @return column index, starting from 1 (inclusive). In case of invalid cell reference, zero.
     */
    public static int getColumnIndexFromCellRef( String cellRef ) {
        int result = 0;
        if( cellRef != null ) {
            for( int index = 0; index < cellRef.length(); index++ ) {
                char c = cellRef.charAt( index );
                if( c < 'A' || c > 'Z' ) {
                    return result;
                }
                result = result * 26 + (c - 'A' + 1);
            }
        }
        return result;
    }

    /** Returns index of first column in sheet (inclusive). Minimum value is 1.
     * @return index of first column in sheet.
     */
    public int getFirstColumnIndex() {
        return firstColumnIndex;
    }

    /** Returns index of last column in sheet (inclusive). Minimum value is 1.
     * @return index of last column in sheet.
     */
    public int getLastColumnIndex() {
        return lastColumnIndex;
    }
}
