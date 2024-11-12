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
package com.inet.excel;

import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Collections;
import java.util.List;

import com.inet.excel.parser.ValueType;

/** Class for result set used to provide information about available procedures and procedure columns.
 */
public class ExcelDatabaseResultSet extends ExcelResultSet {

    private List<List<Object>> rows;
    private int currentRowIndex;
    private boolean wasNull;

    /** Constructor of the class.
     * @param columnNames list of column names.
     * @param rows list of rows with data corresponding to specified columns.
     * @throws IllegalArgumentException if any of given lists is null; if number of values in all rows does not match number of columns.
     */
    public ExcelDatabaseResultSet( List<String> columnNames, List<List<Object>> rows ) {
        super( columnNames );
        if( rows == null ) {
            throw new IllegalArgumentException( "list of rows must not be null" );
        }
        if( rows.stream().anyMatch( row -> row.size() != columnNames.size() ) ) {
            throw new IllegalArgumentException( "number of values in all rows must match number of columns" );
        }
        this.rows = rows;
        this.currentRowIndex = -1;
        this.wasNull = false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean next() throws SQLException {
        throwIfAlreadyClosed();
        currentRowIndex++;
        if( currentRowIndex >= rows.size() ) {
            currentRowIndex = rows.size();
            return false;
        }
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SQLException {
        rows = null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSetMetaData getMetaData() throws SQLException {
        List<String> columnNames = getColumnNames();
        List<ValueType> columnTypes = Collections.nCopies( columnNames.size(), ValueType.VARCHAR );
        return new ExcelSheetResultSetMetaData( "", "", columnNames, columnTypes );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isAfterLast() throws SQLException {
        throwIfAlreadyClosed();
        return currentRowIndex >= rows.size();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getRow() throws SQLException {
        throwIfAlreadyClosed();
        if( currentRowIndex >= rows.size() ) {
            return 0;
        }
        return currentRowIndex + 1;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isClosed() throws SQLException {
        return rows == null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected <T> T getValue( int columnIndex ) throws SQLException {
        throwIfAlreadyClosedOrReachedEnd();
        throwIfColumnIndexIsInvalid( columnIndex );
        T value = (T)rows.get( currentRowIndex ).get( columnIndex - 1 );
        wasNull = value == null;
        return value;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean wasNull() throws SQLException {
        throwIfAlreadyClosed();
        return wasNull;
    }
}
