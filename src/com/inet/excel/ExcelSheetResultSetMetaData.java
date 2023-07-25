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
import java.sql.Types;
import java.util.List;

/** Implementation of {@link ResultSetMetaData} for {@link ExcelSheetResultSet}.
 */
public class ExcelSheetResultSetMetaData implements ResultSetMetaData {

    private String fileName;
    private String sheetName;
    private List<String> columnNames;

    /** Constructor of the class.
     * @param fileName file name of the Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @param columnNames list of column names.
     * @throws IllegalArgumentException if any of given arguments is null.
     */
    public ExcelSheetResultSetMetaData( String fileName, String sheetName, List<String> columnNames ) {
        if( fileName == null ) {
            throw new IllegalArgumentException( "file name must not be null" );
        }
        if( sheetName == null ) {
            throw new IllegalArgumentException( "sheet name must not be null" );
        }
        if( columnNames == null ) {
            throw new IllegalArgumentException( "list of column names must not be null" );
        }
        this.fileName = fileName;
        this.sheetName = sheetName;
        this.columnNames = columnNames;
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
    public int getColumnCount() throws SQLException {
        return columnNames.size();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isAutoIncrement( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isCaseSensitive( int column ) throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isSearchable( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isCurrency( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int isNullable( int column ) throws SQLException {
        return columnNoNulls;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isSigned( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getColumnDisplaySize( int column ) throws SQLException {
        return 32767; // NOTE: value taken from the website "Excel specifications and limits"
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getColumnLabel( int column ) throws SQLException {
        return getColumnName( column );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getColumnName( int column ) throws SQLException {
        return columnNames.get( column - 1 );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSchemaName( int column ) throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getPrecision( int column ) throws SQLException {
        return ExcelDatabaseMetaData.VARCHAR_COLUMN_SIZE_IN_BYTES;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getScale( int column ) throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getTableName( int column ) throws SQLException {
        return sheetName;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getCatalogName( int column ) throws SQLException {
        return fileName;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getColumnType( int column ) throws SQLException {
        return Types.VARCHAR;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getColumnTypeName( int column ) throws SQLException {
        return "VARCHAR";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isReadOnly( int column ) throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isWritable( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isDefinitelyWritable( int column ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getColumnClassName( int column ) throws SQLException {
        return String.class.getName();
    }
}
