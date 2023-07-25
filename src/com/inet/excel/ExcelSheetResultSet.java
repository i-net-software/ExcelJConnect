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

import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.List;

import com.inet.excel.parser.ExcelParser;

/** Class for result set used to retrieve data of the sheet from Excel document.
 */
public class ExcelSheetResultSet extends ExcelResultSet {

    private final ExcelParser parser;
    private final String sheetName;
    private final int maxRowsPerBatch;
    private final ResultSetMetaData metaData;
    private final int rowCount;

    private List<List<String>> rowBatch;
    private int currentRowIndex;
    private int currentBatchIndex;
    private boolean closed;

    /** Constructor of the class.
     * @param parser component responsible for reading data from Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @param maxRowsPerBatch maximum number of rows read at one time.
     * @throws IllegalArgumentException if any of given arguments is null; if max number of rows per batch is not greater than zero.
     */
    public ExcelSheetResultSet( ExcelParser parser, String sheetName, int maxRowsPerBatch ) {
        super( getColumnNames( parser, sheetName ) );
        if( maxRowsPerBatch <= 0 ) {
            throw new IllegalArgumentException( "max number of rows per batch must be greater than zero" );
        }
        this.parser = parser;
        this.sheetName = sheetName;
        this.maxRowsPerBatch = maxRowsPerBatch;
        this.metaData = new ExcelSheetResultSetMetaData( parser.getFileName(), sheetName, getColumnNames() );
        this.rowCount = parser.getRowCount( sheetName );
        this.currentRowIndex = -1;
        this.currentBatchIndex = -1;
        this.closed = false;
    }

    /** Uses given parser to obtain list of column names from specified sheet, but at the very beginning, it performs null-checks.
     * @param parser component responsible for reading data from Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @return list of column names from specified sheet.
     * @throws IllegalArgumentException if any of given arguments is null.
     */
    private static List<String> getColumnNames( ExcelParser parser, String sheetName ) {
        if( parser == null ) {
            throw new IllegalArgumentException( "parser must not be null" );
        }
        if( sheetName == null ) {
            throw new IllegalArgumentException( "sheet name must not be null" );
        }
        return parser.getColumnNames( sheetName );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean next() throws SQLException {
        throwIfAlreadyClosed();

        if( currentRowIndex + 1 >= rowCount ) {
            currentRowIndex = rowCount;
            return false;
        }

        if( currentBatchIndex == -1 || currentBatchIndex == rowBatch.size() - 1 ) {
            int firstRowIndex = currentRowIndex + 2; //NOTE: +1 because we need next element; another +1 because currentRowIndex starts with 0 and indexes required by getRows() start with 1
            int lastRowIndex = firstRowIndex + maxRowsPerBatch - 1; //NOTE: -1 because row specified by lastRowIndex is going to be included in resulting list
            rowBatch = parser.getRows( sheetName, firstRowIndex, Math.min( lastRowIndex, rowCount ) );
            currentBatchIndex = 0;
        } else {
            currentBatchIndex++;
        }
        currentRowIndex++;
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SQLException {
        closed = true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSetMetaData getMetaData() throws SQLException {
        throwIfAlreadyClosed();
        return metaData;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isAfterLast() throws SQLException {
        throwIfAlreadyClosed();
        return currentRowIndex >= rowCount;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getRow() throws SQLException {
        throwIfAlreadyClosed();
        if( currentRowIndex >= rowCount ) {
            return 0;
        }
        return currentRowIndex + 1;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isClosed() throws SQLException {
        return closed;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected <T> T getValue( int columnIndex ) throws SQLException {
        throwIfAlreadyClosedOrReachedEnd();
        throwIfColumnIndexIsInvalid( columnIndex );
        return (T)rowBatch.get( currentBatchIndex ).get( columnIndex - 1 );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean wasNull() throws SQLException {
        throwIfAlreadyClosed();
        return false;
    }
}
