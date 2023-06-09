package com.inet.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.sql.SQLException;
import java.util.Arrays;
import java.util.List;
import java.util.UUID;

import org.junit.jupiter.api.Test;

class ExcelDriverTest {

    @Test
    void acceptsURL_throws_exception_if_url_is_null() {
        ExcelDriver driver = newDriver();
        assertThrows( SQLException.class, () -> driver.acceptsURL( null ) );
    }

    @Test
    void acceptsURL_returns_false_if_url_does_not_start_with_expected_prefix() throws SQLException {
        List<String> invalidURLs = Arrays.asList( "", //
                                                  " ", //
                                                  "   ", //
                                                  " " + ExcelDriver.URL_PREFIX, //
                                                  UUID.randomUUID().toString() + ExcelDriver.URL_PREFIX );

        for( String url : invalidURLs ) {
            assertFalse( newDriver().acceptsURL( url ), url );
        }
    }

    @Test
    void acceptsURL_returns_true_if_url_starts_with_expected_prefix() throws SQLException {
        // method "acceptsURL" checks whether sub protocol matches - it does not validate whole URL
        List<String> validURLs = Arrays.asList( ExcelDriver.URL_PREFIX, //
                                                ExcelDriver.URL_PREFIX + UUID.randomUUID().toString() );

        for( String url : validURLs ) {
            assertTrue( newDriver().acceptsURL( url ), url );
        }
    }

    @Test
    public void getMajorVersion_should_return_current_major_version() {
        assertEquals( 1, newDriver().getMajorVersion() );
    }

    @Test
    public void getMinorVersion_should_return_current_minor_version() {
        assertEquals( 0, newDriver().getMinorVersion() );
    }

    @Test
    public void jdbcCompliant_returns_false() {
        assertFalse( newDriver().jdbcCompliant() );
    }

    private ExcelDriver newDriver() {
        return new ExcelDriver();
    }
}
