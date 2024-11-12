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

import java.io.InputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.Array;
import java.sql.Blob;
import java.sql.CallableStatement;
import java.sql.Clob;
import java.sql.Date;
import java.sql.NClob;
import java.sql.ParameterMetaData;
import java.sql.Ref;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.RowId;
import java.sql.SQLException;
import java.sql.SQLXML;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Calendar;
import java.util.Map;

import com.inet.excel.parser.ExcelParser;

/** Class for statement whose sole purpose is to provide result set with data of the sheet from Excel document.
 */
class ExcelCallableStatement extends ExcelStatement implements CallableStatement {

    private final String sheetName;

    /** Constructor of the class.
     * @param parser component responsible for reading data from Excel document.
     * @param sql the sql to call
     * @throws IllegalArgumentException if any of given arguments is null.
     */
    public ExcelCallableStatement( ExcelParser parser, String sql ) throws SQLException {
        super( parser );
        if( sql == null ) {
            throw new IllegalArgumentException( "sql name must not be null" );
        }
        this.sheetName = getSheetName( sql );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet executeQuery() throws SQLException {
        throwIfAlreadyClosed();
        return new ExcelSheetResultSet( getParser(), sheetName, 50 );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSetMetaData getMetaData() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int executeUpdate() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNull( int parameterIndex, int sqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBoolean( int parameterIndex, boolean x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setByte( int parameterIndex, byte x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setShort( int parameterIndex, short x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setInt( int parameterIndex, int x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setLong( int parameterIndex, long x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setFloat( int parameterIndex, float x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDouble( int parameterIndex, double x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBigDecimal( int parameterIndex, BigDecimal x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setString( int parameterIndex, String x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBytes( int parameterIndex, byte[] x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDate( int parameterIndex, Date x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTime( int parameterIndex, Time x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTimestamp( int parameterIndex, Timestamp x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( int parameterIndex, InputStream x, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setUnicodeStream( int parameterIndex, InputStream x, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( int parameterIndex, InputStream x, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void clearParameters() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( int parameterIndex, Object x, int targetSqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( int parameterIndex, Object x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean execute() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void addBatch() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( int parameterIndex, Reader reader, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setRef( int parameterIndex, Ref x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( int parameterIndex, Blob x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( int parameterIndex, Clob x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setArray( int parameterIndex, Array x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDate( int parameterIndex, Date x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTime( int parameterIndex, Time x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTimestamp( int parameterIndex, Timestamp x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNull( int parameterIndex, int sqlType, String typeName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setURL( int parameterIndex, URL x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ParameterMetaData getParameterMetaData() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setRowId( int parameterIndex, RowId x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNString( int parameterIndex, String value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNCharacterStream( int parameterIndex, Reader value, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( int parameterIndex, NClob value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( int parameterIndex, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( int parameterIndex, InputStream inputStream, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( int parameterIndex, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setSQLXML( int parameterIndex, SQLXML xmlObject ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( int parameterIndex, Object x, int targetSqlType, int scaleOrLength ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( int parameterIndex, InputStream x, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( int parameterIndex, InputStream x, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( int parameterIndex, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( int parameterIndex, InputStream x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( int parameterIndex, InputStream x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( int parameterIndex, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNCharacterStream( int parameterIndex, Reader value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( int parameterIndex, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( int parameterIndex, InputStream inputStream ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( int parameterIndex, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( int parameterIndex, int sqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( int parameterIndex, int sqlType, int scale ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean wasNull() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getString( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean getBoolean( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte getByte( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public short getShort( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getInt( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public long getLong( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public float getFloat( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public double getDouble( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( int parameterIndex, int scale ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte[] getBytes( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( int parameterIndex, Map<String, Class<?>> map ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Ref getRef( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Blob getBlob( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Clob getClob( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Array getArray( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( int parameterIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( int parameterIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( int parameterIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( int parameterIndex, int sqlType, String typeName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( String parameterName, int sqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( String parameterName, int sqlType, int scale ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void registerOutParameter( String parameterName, int sqlType, String typeName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public URL getURL( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setURL( String parameterName, URL val ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNull( String parameterName, int sqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBoolean( String parameterName, boolean x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setByte( String parameterName, byte x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setShort( String parameterName, short x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setInt( String parameterName, int x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setLong( String parameterName, long x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setFloat( String parameterName, float x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDouble( String parameterName, double x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBigDecimal( String parameterName, BigDecimal x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setString( String parameterName, String x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBytes( String parameterName, byte[] x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDate( String parameterName, Date x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTime( String parameterName, Time x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTimestamp( String parameterName, Timestamp x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( String parameterName, InputStream x, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( String parameterName, InputStream x, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( String parameterName, Object x, int targetSqlType, int scale ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( String parameterName, Object x, int targetSqlType ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setObject( String parameterName, Object x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( String parameterName, Reader reader, int length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setDate( String parameterName, Date x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTime( String parameterName, Time x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setTimestamp( String parameterName, Timestamp x, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNull( String parameterName, int sqlType, String typeName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getString( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean getBoolean( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte getByte( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public short getShort( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getInt( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public long getLong( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public float getFloat( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public double getDouble( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte[] getBytes( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( String parameterName, Map<String, Class<?>> map ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Ref getRef( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Blob getBlob( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Clob getClob( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Array getArray( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( String parameterName, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( String parameterName, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( String parameterName, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public URL getURL( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RowId getRowId( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RowId getRowId( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setRowId( String parameterName, RowId x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNString( String parameterName, String value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNCharacterStream( String parameterName, Reader value, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( String parameterName, NClob value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( String parameterName, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( String parameterName, InputStream inputStream, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( String parameterName, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public NClob getNClob( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public NClob getNClob( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setSQLXML( String parameterName, SQLXML xmlObject ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLXML getSQLXML( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLXML getSQLXML( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getNString( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getNString( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getNCharacterStream( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getNCharacterStream( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getCharacterStream( int parameterIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getCharacterStream( String parameterName ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( String parameterName, Blob x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( String parameterName, Clob x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( String parameterName, InputStream x, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( String parameterName, InputStream x, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( String parameterName, Reader reader, long length ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setAsciiStream( String parameterName, InputStream x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBinaryStream( String parameterName, InputStream x ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setCharacterStream( String parameterName, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNCharacterStream( String parameterName, Reader value ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setClob( String parameterName, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setBlob( String parameterName, InputStream inputStream ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setNClob( String parameterName, Reader reader ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();

    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T getObject( int parameterIndex, Class<T> type ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T getObject( String parameterName, Class<T> type ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

}
