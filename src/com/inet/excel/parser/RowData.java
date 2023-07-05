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

import java.util.ArrayList;
import java.util.List;

/** Container for data of cells belonging to single row.
 */
public class RowData {

    private int rowIndex;
    private List<CellData> cellsInRow = new ArrayList<>();

    /** Creates instance representing row with specified index.
     * @param rowIndex index of the row. Minimum value is 1.
     * @throws IllegalArgumentException if specified index is smaller than 1.
     */
    public RowData( int rowIndex ) {
        if( rowIndex < 1 ) {
            throw new IllegalArgumentException( "index must be greater than zero" );
        }
        this.rowIndex = rowIndex;
    }

    /** Returns index of the row. Minimum value is 1.
     * @return index of the row.
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /** Adds data of cell belonging to the row, which is represented by this container.
     * @param cellData data of the cell.
     */
    public void addCellData( CellData cellData ) {
        cellsInRow.add( cellData );
    }

    /** Returns list of all cell data included in this container.
     * @return list of all cell data included in this container.
     */
    public List<CellData> getCellsInRow() {
        return cellsInRow;
    }

    /** Container for data of single cell.
     */
    public static class CellData {
        private String r, s, t, v;

        /** Returns reference to the location of this cell. For example "A1".
         * @return reference to the location of this cell.
         */
        public String getR() {
            return r;
        }

        /** Sets reference to the location of this cell.
         * @param r reference to the location of this cell. For example "A1".
         */
        public void setR( String r ) {
            this.r = r;
        }

        /** Returns index of this cell's style. For example "1".
         * @return index of this cell's style.
         */
        public String getS() {
            return s;
        }

        /** Sets index of this cell's style.
         * @param s index of this cell's style. For example "1".
         */
        public void setS( String s ) {
            this.s = s;
        }

        /** Returns enumeration representing this cell's data type. For example "s" for shared string.
         * @return enumeration representing this cell's data type.
         */
        public String getT() {
            return t;
        }

        /** Sets enumeration representing this cell's data type.
         * @param t enumeration representing this cell's data type. For example "s" for shared string.
         */
        public void setT( String t ) {
            this.t = t;
        }

        /** Returns this cell's value.
         * @return this cell's value.
         */
        public String getV() {
            return v;
        }

        /** Sets this cell's value.
         * @param v this cell's value.
         */
        public void setV( String v ) {
            this.v = v;
        }
    }
}
