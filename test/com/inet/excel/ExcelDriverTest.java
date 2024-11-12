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

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;
import java.util.UUID;

import org.junit.jupiter.api.Test;

class ExcelDriverTest {

    @Test
    public void connect_throws_exception_if_url_is_null() {
        ExcelDriver driver = newDriver();
        assertThrows( SQLException.class, () -> driver.connect( null, new Properties() ) );
    }

    @Test
    public void connect_returns_null_if_url_does_not_start_with_expected_prefix() throws SQLException {
        for( String url : getURLsWhichDoNotStartWithExpectedPrefix() ) {
            assertNull( newDriver().connect( url, new Properties() ), url );
        }
    }

    @Test
    public void connect_throws_exception_if_path_to_excel_file_is_not_specified_in_url() {
        List<String> missingFilePathURLs = Arrays.asList( ExcelDriver.URL_PREFIX, //
                                                          ExcelDriver.URL_PREFIX + "     ", //
                                                          ExcelDriver.URL_PREFIX + "?", //
                                                          ExcelDriver.URL_PREFIX + "?hasHeaderRow=false", //
                                                          ExcelDriver.URL_PREFIX + "    ?", //
                                                          ExcelDriver.URL_PREFIX + "    ?hasHeaderRow=false" );

        for( String url : missingFilePathURLs ) {
            ExcelDriver driver = newDriver();
            assertThrows( SQLException.class, () -> driver.connect( url, new Properties() ), url );
        }
    }

    @Test
    public void connect_throws_exception_if_path_to_excel_file_does_not_represent_existing_file() throws IOException {
        ExcelDriver driver = newDriver();

        Path folder = Files.createTempDirectory( "ExcelDriverTest_" + UUID.randomUUID() );
        assertTrue( Files.isDirectory( folder ) ); // precondition check
        String urlWithPathToExistingFolder = ExcelDriver.URL_PREFIX + folder.toAbsolutePath().toString();
        assertThrows( SQLException.class, () -> driver.connect( urlWithPathToExistingFolder, new Properties() ) );

        Path nonExistingFile = folder.getParent().resolve( "nonExistingFile.xlsx" );
        assertFalse( Files.exists( nonExistingFile ) ); // precondition check
        String urlWithPathToNonExistingFile = ExcelDriver.URL_PREFIX + nonExistingFile.toAbsolutePath().toString();
        assertThrows( SQLException.class, () -> driver.connect( urlWithPathToNonExistingFile, new Properties() ) );
    }

    @Test
    public void connect_returns_connection_to_specified_existing_file() throws IOException, SQLException {
        Path file = Files.createTempFile(  "ExcelDriverTest_" + UUID.randomUUID(), ".xlsx" );
        assertTrue( Files.isRegularFile( file ) ); // precondition check
        String fileAbsolutePath = file.toAbsolutePath().toString();

        String urlWithPathToExistingFile = ExcelDriver.URL_PREFIX + fileAbsolutePath;
        assertNotNull( newDriver().connect( urlWithPathToExistingFile, new Properties() ) );

        urlWithPathToExistingFile = ExcelDriver.URL_PREFIX + "file:///" + fileAbsolutePath;
        assertNotNull( newDriver().connect( urlWithPathToExistingFile, new Properties() ) );

        urlWithPathToExistingFile = ExcelDriver.URL_PREFIX + "file://localhost/" + fileAbsolutePath;
        assertNotNull( newDriver().connect( urlWithPathToExistingFile, new Properties() ) );
    }

    @Test
    public void acceptsURL_throws_exception_if_url_is_null() {
        ExcelDriver driver = newDriver();
        assertThrows( SQLException.class, () -> driver.acceptsURL( null ) );
    }

    @Test
    public void acceptsURL_returns_false_if_url_does_not_start_with_expected_prefix() throws SQLException {
        for( String url : getURLsWhichDoNotStartWithExpectedPrefix() ) {
            assertFalse( newDriver().acceptsURL( url ), url );
        }
    }

    @Test
    public void acceptsURL_returns_true_if_url_starts_with_expected_prefix() throws SQLException {
        // method "acceptsURL" checks whether sub protocol matches - it does not validate whole URL
        List<String> validURLs = Arrays.asList( ExcelDriver.URL_PREFIX, //
                                                ExcelDriver.URL_PREFIX + UUID.randomUUID().toString() );

        for( String url : validURLs ) {
            assertTrue( newDriver().acceptsURL( url ), url );
        }
    }

    @Test
    public void getMajorVersion_should_return_current_major_version() {
        assertEquals( ExcelDriver.MAJOR_VERSION, newDriver().getMajorVersion() );
    }

    @Test
    public void getMinorVersion_should_return_current_minor_version() {
        assertEquals( ExcelDriver.MINOR_VERSION, newDriver().getMinorVersion() );
    }

    @Test
    public void jdbcCompliant_returns_false() {
        assertFalse( newDriver().jdbcCompliant() );
    }

    private ExcelDriver newDriver() {
        return new ExcelDriver();
    }

    private List<String> getURLsWhichDoNotStartWithExpectedPrefix() {
        return Arrays.asList( "", //
                              " ", //
                              "   ", //
                              " " + ExcelDriver.URL_PREFIX, //
                              UUID.randomUUID().toString() + ExcelDriver.URL_PREFIX );
    }
}
