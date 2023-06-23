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

public class ExcelDatabaseResultSet extends ExcelResultSet {

    private List<List<Object>> rows;
    private int currentRowIndex;

    public ExcelDatabaseResultSet( List<String> columnNames, List<List<Object>> rows ) {
        super( columnNames );
        this.rows = rows;
        this.currentRowIndex = -1;
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
        // TODO Auto-generated method stub
        return null;
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
        return (T)rows.get( currentRowIndex ).get( columnIndex - 1 );
    }
}
