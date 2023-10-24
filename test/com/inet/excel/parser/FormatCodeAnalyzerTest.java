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
package com.inet.excel.parser;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;

public class FormatCodeAnalyzerTest {

    private static final String QUOTE            = "\"";
    private static final String BACKSLASH        = "\\";
    private static final String DOUBLE_BACKSLASH = "\\\\";

    @Test
    public void varchar() {
        List<String> cases = Arrays.asList( "", //
                                            "    ", //
                                            BACKSLASH + "d", //
                                            BACKSLASH + "m", //
                                            BACKSLASH + "y", //
                                            DOUBLE_BACKSLASH + BACKSLASH + "d", //
                                            DOUBLE_BACKSLASH + BACKSLASH + "m", //
                                            DOUBLE_BACKSLASH + BACKSLASH + "y", //
                                            QUOTE + BACKSLASH + "yy" + QUOTE, //
                                            QUOTE + BACKSLASH + "mm" + QUOTE, //
                                            QUOTE + BACKSLASH + "dd" + QUOTE );

        for( String formatCode : cases ) {
            assertVarchar( formatCode.toLowerCase() );
            assertVarchar( formatCode.toUpperCase() );
        }
        assertVarchar( null );
    }

    @Test
    public void date() {
        List<String> cases = Arrays.asList( "d", //
                                            "m", //
                                            ":m", //
                                            "m:", //
                                            ":m:", //
                                            "y", //
                                            "m/d/y", //
                                            "d-mmm", //
                                            "mmm-yy", //
                                            "m/d", //
                                            "dd/mm/yyyy", //
                                            "yyyy-mm-dd", //
                                            "mmmm d, yyyy", //
                                            "mmmmm", //
                                            "dddd, mmmm d, yyyy", //
                                            DOUBLE_BACKSLASH + "d", //
                                            DOUBLE_BACKSLASH + "m", //
                                            DOUBLE_BACKSLASH + "y", //
                                            DOUBLE_BACKSLASH + DOUBLE_BACKSLASH + "d", //
                                            DOUBLE_BACKSLASH + DOUBLE_BACKSLASH + "m", //
                                            DOUBLE_BACKSLASH + DOUBLE_BACKSLASH + "y", //
                                            BACKSLASH + "yy", //
                                            BACKSLASH + "mm", //
                                            BACKSLASH + "dd" );

        for( String formatCode : cases ) {
            assertDate( formatCode.toLowerCase() );
            assertDate( formatCode.toUpperCase() );
        }
    }

    @Test
    public void time() {
        List<String> cases = Arrays.asList( "h", //
                                            "s", //
                                            "h:mm", //
                                            "h:mm:ss", //
                                            "h:mm AM/PM", //
                                            "mm:ss", //
                                            "h:m", //
                                            "h:s", //
                                            "m:s" );

        for( String formatCode : cases ) {
            assertTime( formatCode.toLowerCase() );
            assertTime( formatCode.toUpperCase() );
        }
    }

    @Test
    public void timestamp() {
        List<String> cases = Arrays.asList( "m/d/yyyy h:mm", //
                                            "m/d/yy h:mm AM/PM", //
                                            "hh:mm:ss dddd/mmmm/yyyy" );
        for( String formatCode : cases ) {
            assertTimestamp( formatCode.toLowerCase() );
            assertTimestamp( formatCode.toUpperCase() );
        }
    }

    private void assertVarchar( String formatCode ) {
        assertValueType( ValueType.VARCHAR, formatCode );
    }

    private void assertTime( String formatCode ) {
        assertValueType( ValueType.TIME, formatCode );
    }

    private void assertDate( String formatCode ) {
        assertValueType( ValueType.DATE, formatCode );
    }

    private void assertTimestamp( String formatCode ) {
        assertValueType( ValueType.TIMESTAMP, formatCode );
    }

    private void assertValueType( ValueType type, String formatCode ) {
        String msg = "formatCode=\"" + formatCode + "\"";
        assertEquals( type, FormatCodeAnalyzer.recognizeValueType( formatCode ), msg );
    }
}
