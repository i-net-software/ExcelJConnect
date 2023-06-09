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

public class ExcelResultSetMetaData implements ResultSetMetaData {

    private List<String> columnNames;

    public ExcelResultSetMetaData( List<String> columnNames ) {
        //TODO ensure list is not null
        this.columnNames = columnNames;
    }

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
    public int getColumnCount() throws SQLException {
        // TODO Auto-generated method stub
        return columnNames.size();
    }

    @Override
    public boolean isAutoIncrement( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean isCaseSensitive( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean isSearchable( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean isCurrency( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public int isNullable( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public boolean isSigned( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public int getColumnDisplaySize( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public String getColumnLabel( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public String getColumnName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return columnNames.get( column );
    }

    @Override
    public String getSchemaName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public int getPrecision( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public int getScale( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public String getTableName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public String getCatalogName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public int getColumnType( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public String getColumnTypeName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

    @Override
    public boolean isReadOnly( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean isWritable( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public boolean isDefinitelyWritable( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return false;
    }

    @Override
    public String getColumnClassName( int column ) throws SQLException {
        // TODO Auto-generated method stub
        return null;
    }

}
