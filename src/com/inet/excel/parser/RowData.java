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
