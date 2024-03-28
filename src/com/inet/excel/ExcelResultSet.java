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

import java.io.InputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.Array;
import java.sql.Blob;
import java.sql.Clob;
import java.sql.Date;
import java.sql.NClob;
import java.sql.Ref;
import java.sql.ResultSet;
import java.sql.RowId;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.SQLXML;
import java.sql.Statement;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Base class for result sets used within Excel Driver project.
 */
public abstract class ExcelResultSet implements ResultSet {

    private final List<String> columnNames;

    /** Constructor of the class.
     * @param columnNames list of column names.
     * @throws IllegalArgumentException if given list is null.
     */
    public ExcelResultSet( List<String> columnNames ) {
        if( columnNames == null ) {
            throw new IllegalArgumentException( "list of column names must not be null" );
        }
        this.columnNames = Collections.unmodifiableList( columnNames );
    }

    /** Retrieves the value of the designated column in the current row of this ResultSet object.
     * @param <T> type of the returned value.
     * @param columnIndex the first column is 1, the second is 2, ...
     * @return the column value.
     * @throws SQLException if the columnIndex is not valid; if a database access error occurs or this method is called on a closed result set.
     */
    abstract protected <T> T getValue( int columnIndex ) throws SQLException;

    /** Throws exception if result set is already closed.
     * @throws SQLException if result set is already closed.
     */
    protected void throwIfAlreadyClosed() throws SQLException {
        if( isClosed() ) {
            throw new SQLException( "ResultSet: already closed" );
        }
    }

    /** Throws exception if result set is closed or its cursor is after the last row.
     * @throws SQLException if result set is closed or its cursor is after the last row.
     */
    protected void throwIfAlreadyClosedOrReachedEnd() throws SQLException {
        throwIfAlreadyClosed();
        if( isAfterLast() ) {
            throw new SQLException( "ResultSet: already reached end" );
        }
    }

    /** Throws exception if the columnIndex is not valid.
     * @param columnIndex the first column is 1, the second is 2, ...
     * @throws SQLException if the columnIndex is not valid.
     */
    protected void throwIfColumnIndexIsInvalid( int columnIndex ) throws SQLException {
        if( columnIndex < 1 || columnIndex > columnNames.size() ) {
            throw new SQLException( "ResultSet: invalid column index " + columnIndex + " for column count " + columnNames.size() );
        }
    }

    /** Throws exception indicating that data may not be updated using this result set due to its concurrency mode.
     * @throws SQLException exception indicating that data may not be updated using this result set due to its concurrency mode.
     */
    protected void throwExceptionDueToConcurrencyMode() throws SQLException {
        throw new SQLException( "Data may not be updated using this result set" );
    }

    /** Throws exception indicating that operation is not allowed on result set of type {@link ResultSet#TYPE_FORWARD_ONLY}.
     * @throws SQLException exception indicating that operation is not allowed on result set of type {@link ResultSet#TYPE_FORWARD_ONLY}.
     */
    protected void throwExceptionDueToResultSetType() throws SQLException {
        throw new SQLException( "Operation is not allowed on result set of type \"TYPE_FORWARD_ONLY\"" );
    }

    /** Returns list of column names.
     * @return list of column names.
     */
    protected List<String> getColumnNames() {
        return columnNames;
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
    public String getString( int columnIndex ) throws SQLException {
        return getValue( columnIndex  );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean getBoolean( int columnIndex ) throws SQLException {
        Boolean value = getValue( columnIndex );
        if( value == null ) {
            return false;
        }
        return value.booleanValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte getByte( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.byteValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public short getShort( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.shortValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getInt( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.intValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public long getLong( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.longValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public float getFloat( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.floatValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public double getDouble( int columnIndex ) throws SQLException {
        Number value = getValue( columnIndex );
        if( value == null ) {
            return 0;
        }
        return value.doubleValue();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( int columnIndex, int scale ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte[] getBytes( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getAsciiStream( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getUnicodeStream( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getBinaryStream( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getString( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getString( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean getBoolean( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getBoolean( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte getByte( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getByte( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public short getShort( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getShort( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getInt( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getInt( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public long getLong( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getLong( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public float getFloat( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getFloat( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public double getDouble( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getDouble( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( String columnLabel, int scale ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public byte[] getBytes( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getBytes( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getDate( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getTime( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getTimestamp( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getAsciiStream( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getUnicodeStream( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public InputStream getBinaryStream( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLWarning getWarnings() throws SQLException {
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void clearWarnings() throws SQLException {
        // nothing to do
    }

    @Override
    public String getCursorName() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getObject( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int findColumn( String columnLabel ) throws SQLException {
        throwIfAlreadyClosed();
        for( int index = 0; index < columnNames.size(); index++ ) {
            if( Objects.equals( columnLabel, columnNames.get( index ) ) ) {
                return index + 1;
            }
        }
        throw new SQLException( "ResultSet: unknown column \"" + columnLabel + "\"." );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getCharacterStream( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getCharacterStream( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getCharacterStream( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( int columnIndex ) throws SQLException {
        return getValue( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public BigDecimal getBigDecimal( String columnLabel ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getBigDecimal( columnIndex );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isBeforeFirst() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isFirst() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isLast() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void beforeFirst() throws SQLException {
        throwExceptionDueToResultSetType();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void afterLast() throws SQLException {
        throwExceptionDueToResultSetType();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean first() throws SQLException {
        throwExceptionDueToResultSetType();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean last() throws SQLException {
        throwExceptionDueToResultSetType();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean absolute( int row ) throws SQLException {
        throwExceptionDueToResultSetType();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean relative( int rows ) throws SQLException {
        throwExceptionDueToResultSetType();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean previous() throws SQLException {
        throwExceptionDueToResultSetType();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setFetchDirection( int direction ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getFetchDirection() throws SQLException {
        return ResultSet.FETCH_FORWARD;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setFetchSize( int rows ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getFetchSize() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getType() throws SQLException {
        return ResultSet.TYPE_FORWARD_ONLY;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getConcurrency() throws SQLException {
        return CONCUR_READ_ONLY;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean rowUpdated() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean rowInserted() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean rowDeleted() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNull( int columnIndex ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBoolean( int columnIndex, boolean x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateByte( int columnIndex, byte x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateShort( int columnIndex, short x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateInt( int columnIndex, int x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateLong( int columnIndex, long x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateFloat( int columnIndex, float x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateDouble( int columnIndex, double x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBigDecimal( int columnIndex, BigDecimal x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateString( int columnIndex, String x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBytes( int columnIndex, byte[] x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateDate( int columnIndex, Date x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateTime( int columnIndex, Time x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateTimestamp( int columnIndex, Timestamp x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( int columnIndex, InputStream x, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( int columnIndex, InputStream x, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( int columnIndex, Reader x, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateObject( int columnIndex, Object x, int scaleOrLength ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateObject( int columnIndex, Object x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNull( String columnLabel ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBoolean( String columnLabel, boolean x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateByte( String columnLabel, byte x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateShort( String columnLabel, short x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateInt( String columnLabel, int x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateLong( String columnLabel, long x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateFloat( String columnLabel, float x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateDouble( String columnLabel, double x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBigDecimal( String columnLabel, BigDecimal x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateString( String columnLabel, String x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBytes( String columnLabel, byte[] x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateDate( String columnLabel, Date x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateTime( String columnLabel, Time x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateTimestamp( String columnLabel, Timestamp x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( String columnLabel, InputStream x, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( String columnLabel, InputStream x, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( String columnLabel, Reader reader, int length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateObject( String columnLabel, Object x, int scaleOrLength ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateObject( String columnLabel, Object x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void insertRow() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateRow() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void deleteRow() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void refreshRow() throws SQLException {
        throwExceptionDueToResultSetType();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void cancelRowUpdates() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void moveToInsertRow() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void moveToCurrentRow() throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Statement getStatement() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( int columnIndex, Map<String, Class<?>> map ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Ref getRef( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Blob getBlob( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Clob getClob( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Array getArray( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Object getObject( String columnLabel, Map<String, Class<?>> map ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Ref getRef( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Blob getBlob( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Clob getClob( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Array getArray( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( int columnIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Date getDate( String columnLabel, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( int columnIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Time getTime( String columnLabel, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( int columnIndex, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Timestamp getTimestamp( String columnLabel, Calendar cal ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public URL getURL( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public URL getURL( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateRef( int columnIndex, Ref x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateRef( String columnLabel, Ref x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( int columnIndex, Blob x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( String columnLabel, Blob x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( int columnIndex, Clob x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( String columnLabel, Clob x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateArray( int columnIndex, Array x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateArray( String columnLabel, Array x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RowId getRowId( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RowId getRowId( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateRowId( int columnIndex, RowId x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateRowId( String columnLabel, RowId x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getHoldability() throws SQLException {
        throwIfAlreadyClosed();
        return ResultSet.HOLD_CURSORS_OVER_COMMIT;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNString( int columnIndex, String nString ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNString( String columnLabel, String nString ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( int columnIndex, NClob nClob ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( String columnLabel, NClob nClob ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public NClob getNClob( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public NClob getNClob( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLXML getSQLXML( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public SQLXML getSQLXML( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateSQLXML( int columnIndex, SQLXML xmlObject ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateSQLXML( String columnLabel, SQLXML xmlObject ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getNString( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getNString( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getNCharacterStream( int columnIndex ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Reader getNCharacterStream( String columnLabel ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNCharacterStream( int columnIndex, Reader x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNCharacterStream( String columnLabel, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( int columnIndex, InputStream x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( int columnIndex, InputStream x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( int columnIndex, Reader x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( String columnLabel, InputStream x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( String columnLabel, InputStream x, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( String columnLabel, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( int columnIndex, InputStream inputStream, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( String columnLabel, InputStream inputStream, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( int columnIndex, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( String columnLabel, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( int columnIndex, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( String columnLabel, Reader reader, long length ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNCharacterStream( int columnIndex, Reader x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNCharacterStream( String columnLabel, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( int columnIndex, InputStream x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( int columnIndex, InputStream x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( int columnIndex, Reader x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateAsciiStream( String columnLabel, InputStream x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBinaryStream( String columnLabel, InputStream x ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateCharacterStream( String columnLabel, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( int columnIndex, InputStream inputStream ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateBlob( String columnLabel, InputStream inputStream ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( int columnIndex, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateClob( String columnLabel, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( int columnIndex, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void updateNClob( String columnLabel, Reader reader ) throws SQLException {
        throwExceptionDueToConcurrencyMode();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T getObject( int columnIndex, Class<T> type ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T getObject( String columnLabel, Class<T> type ) throws SQLException {
        int columnIndex = findColumn( columnLabel );
        return getObject( columnIndex, type );
    }
}
