package com.inet.excel;

import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverPropertyInfo;
import java.sql.SQLException;
import java.sql.SQLFeatureNotSupportedException;
import java.util.Properties;
import java.util.logging.Logger;

import com.inet.excel.parser.ExcelParser;

public class ExcelDriver implements Driver {

    public static final String URL_PREFIX = "jdbc:inetexcel:";

    /**
     * {@inheritDoc}
     */
    @Override
    public Connection connect( String url, Properties info ) throws SQLException {
        // TODO check whether url is accepted
        ExcelParser parser = new ExcelParser( Paths.get( "D:\\excel_driver_tests\\test.xlsx" ) ); //TODO
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
        // TODO Auto-generated method stub
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMajorVersion() {
        return 1;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMinorVersion() {
        return 0;
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
        // TODO Auto-generated method stub
        return null;
    }
}
