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

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverPropertyInfo;
import java.sql.SQLException;
import java.sql.SQLFeatureNotSupportedException;
import java.util.Properties;
import java.util.logging.Logger;

import com.inet.excel.parser.ExcelParser;

/** Implementation of JDBC Driver, which allows to read data from Excel documents.
 */
public class ExcelDriver implements Driver {

    public static final String URL_PREFIX    = "jdbc:inetexcel:";
    public static final String DRIVER_NAME   = "inetexcel";
    public static final int    MAJOR_VERSION = 1;
    public static final int    MINOR_VERSION = 1;

    /** Throws exception indicating that requested operation is not supported.
     * @throws SQLException exception indicating that requested operation is not supported.
     */
    static void throwExceptionAboutUnsupportedOperation() throws SQLException {
        throw new SQLException( "Unsupported operation" );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Connection connect( String url, Properties info ) throws SQLException {
        if( !acceptsURL( url ) ) {
            return null;
        }

        url = url.substring( URL_PREFIX.length() );

        int questionMarkIndex = url.indexOf( '?' );
        if( questionMarkIndex == 0 ) {
            throw new SQLException( "Excel file is not specified" );
        }

        String filePath = null;
        boolean hasHeaderRow = true;

        if( questionMarkIndex == -1 ) {
            filePath = url;
        } else {
            filePath = url.substring( 0, questionMarkIndex );
            String propertiesPart = url.substring( questionMarkIndex + 1 );
            String[] properties = propertiesPart.split( "&" );

            for( String property : properties ) {
                if( "hasHeaderRow=false".equalsIgnoreCase( property ) ) {
                    hasHeaderRow = false;
                    break;
                }
            }
        }

        if( filePath.trim().isEmpty() ) {
            throw new SQLException( "Excel file is not specified" );
        }

        Path file = Paths.get( filePath ).toAbsolutePath();
        if( !Files.isRegularFile( file ) ) {
            throw new SQLException( "Specified Excel file does not exist" );
        }

        ExcelParser parser = new ExcelParser( file, hasHeaderRow );
        return new ExcelConnection( parser );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean acceptsURL( String url ) throws SQLException {
        if( url == null ) {
            throw new SQLException( "URL must not be null" );
        }
        return url.startsWith( URL_PREFIX );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DriverPropertyInfo[] getPropertyInfo( String url, Properties info ) throws SQLException {
        return new DriverPropertyInfo[0];
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMajorVersion() {
        return MAJOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMinorVersion() {
        return MINOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean jdbcCompliant() {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Logger getParentLogger() throws SQLFeatureNotSupportedException {
        throw new SQLFeatureNotSupportedException();
    }
}
