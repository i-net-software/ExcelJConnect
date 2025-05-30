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

import java.sql.Array;
import java.sql.Blob;
import java.sql.CallableStatement;
import java.sql.Clob;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.NClob;
import java.sql.PreparedStatement;
import java.sql.SQLClientInfoException;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.SQLXML;
import java.sql.Savepoint;
import java.sql.Statement;
import java.sql.Struct;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.Executor;

import com.inet.excel.parser.ExcelParser;

/** Connection intended to be returned by {@link ExcelDriver} in order to read data from Excel documents.
 */
public class ExcelConnection implements Connection {

    private final ExcelParser parser;
    private boolean closed;
    private Runnable onConnectionClose;

    /** Constructor of the class.
     * @param parser component responsible for reading data from Excel document.
     * @param onConnectionClose optional runnable to be executed on connection close.
     * @throws IllegalArgumentException if given parser is null.
     */
    public ExcelConnection( ExcelParser parser, Runnable onConnectionClose ) {
        if( parser == null ) {
            throw new IllegalArgumentException( "parser must not be null" );
        }
        this.parser = parser;
        this.onConnectionClose = onConnectionClose;
        this.closed = false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DatabaseMetaData getMetaData() throws SQLException {
        return new ExcelDatabaseMetaData( parser );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T unwrap( Class<T> iface ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isWrapperFor( Class<?> iface ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Statement createStatement() throws SQLException {
        return new ExcelStatement( parser );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public CallableStatement prepareCall( String sql ) throws SQLException {
        throwIfAlreadyClosed();
        return new ExcelCallableStatement( parser, sql );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String nativeSQL( String sql ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAutoCommit( boolean autoCommit ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean getAutoCommit() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void commit() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void rollback() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws SQLException {
        if( !closed && onConnectionClose != null ) {
            onConnectionClose.run();
        }
        closed = true;
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
    public void setReadOnly( boolean readOnly ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isReadOnly() throws SQLException {
        throwIfAlreadyClosed();
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCatalog( String catalog ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getCatalog() throws SQLException {
        return parser.getFileName();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTransactionIsolation( int level ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getTransactionIsolation() throws SQLException {
        throwIfAlreadyClosed();
        return Connection.TRANSACTION_NONE;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLWarning getWarnings() throws SQLException {
        throwIfAlreadyClosed();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void clearWarnings() throws SQLException {
        throwIfAlreadyClosed();
        // nothing to do
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Statement createStatement( int resultSetType, int resultSetConcurrency ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql, int resultSetType, int resultSetConcurrency ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public CallableStatement prepareCall( String sql, int resultSetType, int resultSetConcurrency ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Map<String, Class<?>> getTypeMap() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTypeMap( Map<String, Class<?>> map ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setHoldability( int holdability ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getHoldability() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Savepoint setSavepoint() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Savepoint setSavepoint( String name ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void rollback( Savepoint savepoint ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void releaseSavepoint( Savepoint savepoint ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Statement createStatement( int resultSetType, int resultSetConcurrency, int resultSetHoldability ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql, int resultSetType, int resultSetConcurrency, int resultSetHoldability ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public CallableStatement prepareCall( String sql, int resultSetType, int resultSetConcurrency, int resultSetHoldability ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql, int autoGeneratedKeys ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql, int[] columnIndexes ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public PreparedStatement prepareStatement( String sql, String[] columnNames ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Clob createClob() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Blob createBlob() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public NClob createNClob() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLXML createSQLXML() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isValid( int timeout ) throws SQLException {
        return !isClosed();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClientInfo( String name, String value ) throws SQLClientInfoException {
        throw new SQLClientInfoException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClientInfo( Properties properties ) throws SQLClientInfoException {
        throw new SQLClientInfoException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getClientInfo( String name ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Properties getClientInfo() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Array createArrayOf( String typeName, Object[] elements ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Struct createStruct( String typeName, Object[] attributes ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setSchema( String schema ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSchema() throws SQLException {
        throwIfAlreadyClosed();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void abort( Executor executor ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNetworkTimeout( Executor executor, int milliseconds ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getNetworkTimeout() throws SQLException {
        throwIfAlreadyClosed();
        return 0;
    }

    /** Throws exception if connection is already closed.
     * @throws SQLException if connection is already closed.
     */
    private void throwIfAlreadyClosed() throws SQLException {
        if( isClosed() ) {
            throw new SQLException( "Connection: already closed" );
        }
    }
}
