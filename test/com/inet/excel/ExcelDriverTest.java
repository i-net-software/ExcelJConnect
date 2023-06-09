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
