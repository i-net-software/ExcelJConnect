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

public class RowData {

    private List<CellData> cellsInRow = new ArrayList<>();

    public void addCellData( CellData cellData ) {
        cellsInRow.add( cellData );
    }

    public List<CellData> getCellsInRow() {
        return cellsInRow;
    }

    public static class CellData {
        private String r, s, t, v;

        public String getR() {
            return r;
        }

        public void setR( String r ) {
            this.r = r;
        }

        public String getS() {
            return s;
        }

        public void setS( String s ) {
            this.s = s;
        }

        public String getT() {
            return t;
        }

        public void setT( String t ) {
            this.t = t;
        }

        public String getV() {
            return v;
        }

        public void setV( String v ) {
            this.v = v;
        }
    }
}
