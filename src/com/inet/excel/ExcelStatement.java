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

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.Statement;

public class ExcelStatement implements Statement {

    @Override
    public <T> T unwrap( Class<T> iface ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public boolean isWrapperFor( Class<?> iface ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public ResultSet executeQuery( String sql ) throws SQLException {
        // TODO Auto-generated method stub
        System.out.println( "executeQuery: " + sql );
        throw new SQLException(); //TODO
    }

    @Override
    public int executeUpdate( String sql ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void close() throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int getMaxFieldSize() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void setMaxFieldSize( int max ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int getMaxRows() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void setMaxRows( int max ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public void setEscapeProcessing( boolean enable ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int getQueryTimeout() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void setQueryTimeout( int seconds ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public void cancel() throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public SQLWarning getWarnings() throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public void clearWarnings() throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public void setCursorName( String name ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public boolean execute( String sql ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public ResultSet getResultSet() throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public int getUpdateCount() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public boolean getMoreResults() throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public void setFetchDirection( int direction ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int getFetchDirection() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void setFetchSize( int rows ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int getFetchSize() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public int getResultSetConcurrency() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public int getResultSetType() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void addBatch( String sql ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public void clearBatch() throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public int[] executeBatch() throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public Connection getConnection() throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public boolean getMoreResults( int current ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public ResultSet getGeneratedKeys() throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public int executeUpdate( String sql, int autoGeneratedKeys ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public int executeUpdate( String sql, int[] columnIndexes ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public int executeUpdate( String sql, String[] columnNames ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public boolean execute( String sql, int autoGeneratedKeys ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean execute( String sql, int[] columnIndexes ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean execute( String sql, String[] columnNames ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public int getResultSetHoldability() throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public boolean isClosed() throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public void setPoolable( boolean poolable ) throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public boolean isPoolable() throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public void closeOnCompletion() throws SQLException {
        // TODO Auto-generated method stub

    }

    @Override
    public boolean isCloseOnCompletion() throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

}
