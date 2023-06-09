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

/** Mutable container for information about row span in single sheet from Excel workbook.
 * It is intended to be used, while parsing sheet data, in order to determine dimension of Excel sheet.
 */
public class RowSpanData {

    private int firstColumnIndex = 0;
    private int lastColumnIndex = 0;

    /** Updates indexes of columns, which constitute row span, based on provided span information.
     * It ignores provided information if it is invalid.
     * Examples of valid span information: "2:2", "4:7", "23:45".
     * Examples of invalid span information: null, "", "   ", "5", ":6", "6:", "A:4", "4:Z", "5:4".
     * @param spans span information to retrieve indexes of columns from.
     */
    public void addSpanRange( String spans ) {
        if( spans == null ) {
            return;
        }

        int spanStart, spanEnd;
        try {
            String[] range = spans.split( ":" );
            spanStart = Integer.parseInt( range[0] );
            spanEnd = Integer.parseInt( range[1] );
        } catch( Exception ex ) {
            return; // ignores data if invalid
        }
        updateIndexes( spanStart, spanEnd );
    }

    /** Updates indexes of columns, which constitute row span, based on provided cell reference.
     * It ignores provided reference if it is invalid.
     * Examples of valid cell references: "A1", "A7", "G14", "BB44", "ASD1".
     * Examples of invalid cell references: null, "", "   ", "5", "$".
     * @param ref cell reference to retrieve index of column from.
     */
    public void addCellRef( String ref ) {
        if( ref == null ) {
            return;
        }
        int columnIndex = SheetDimension.getColumnIndexFromCellRef( ref );
        updateIndexes( columnIndex, columnIndex );
    }

    /** Updates indexes of columns, which constitute row span.
     * It ignores provided data if any of indexes is smaller than 1 or start index is greater than end index.
     * @param spanStart index of first column included in row span. Minimum value is 1.
     * @param spanEnd index of last column included in row span. Minimum value is 1.
     */
    private void updateIndexes( int spanStart, int spanEnd ) {
        if( spanStart < 1 || spanEnd < 1 || spanStart > spanEnd ) {
            return; // ignores data if invalid
        }

        if( firstColumnIndex < 1 || spanStart < firstColumnIndex ) {
            firstColumnIndex = spanStart;
        }
        if( spanEnd > lastColumnIndex ) {
            lastColumnIndex = spanEnd;
        }
    }

    /** Returns true, if container has not been initialized with valid data yet. Otherwise, returns false.
     * @return true, if container has not been initialized with valid data yet. Otherwise, returns false.
     */
    public boolean isEmpty() {
        return firstColumnIndex < 1 || lastColumnIndex < 1;
    }

    /** Returns index of first column included in row span (minimum value is 1). If this method returns 0, it means that container has not been initialized yet and represents no row span.
     * @return index of first column included in row span (minimum value is 1) or 0, if container has not been initialized yet and represents no row span.
     */
    public int getFirstColumnIndex() {
        return firstColumnIndex;
    }

    /** Returns index of last column included in row span (minimum value is 1). If this method returns 0, it means that container has not been initialized yet and represents no row span.
     * @return index of last column included in row span (minimum value is 1) or 0, if container has not been initialized yet and represents no row span.
     */
    public int getLastColumnIndex() {
        return lastColumnIndex;
    }
}
