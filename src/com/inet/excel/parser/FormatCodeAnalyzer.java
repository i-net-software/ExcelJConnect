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

/** Provides method to analyze cell's format code, in order to recognize cell value's type.
 */
public class FormatCodeAnalyzer {

    /** Analyze cell's format code, in order to recognize cell value's type
     * @param formatCode format code to analyze. May be null or empty string.
     * @return cell value's type. Never null.
     */
    public static ValueType recognizeValueType( String formatCode ) {
        if( formatCode == null || formatCode.trim().isEmpty() ) {
            return ValueType.VARCHAR;
        }

        formatCode = formatCode.toLowerCase();

        if( countQuoteCharacters( formatCode ) % 2 != 0 ) {
            // invalid format code
            return ValueType.VARCHAR;
        }

        boolean containsTime = containsTime( formatCode );
        boolean containsDate = containsDate( formatCode );

        if( containsTime && containsDate ) {
            return ValueType.TIMESTAMP;
        }
        if( containsTime ) {
            return ValueType.TIME;
        }
        if( containsDate ) {
            return ValueType.DATE;
        }
        return ValueType.VARCHAR;
    }

    /** Checks whether given format code contains section representing time.
     * @param formatCode format code to analyze. May not be null.
     * @return whether given format code contains section representing time.
     */
    private static boolean containsTime( String formatCode ) {
        if( containsUnescapedSubstring( formatCode, "h" ) ) {
            return true;
        }
        if( containsUnescapedSubstring( formatCode, "h:m" ) ) {
            return true;
        }
        if( containsUnescapedSubstring( formatCode, "m:s" ) ) {
            return true;
        }
        if( containsUnescapedSubstring( formatCode, "s" ) ) {
            return true;
        }
        return false;
    }

    /** Checks whether given format code contains section representing date.
     * @param formatCode format code to analyze. May not be null.
     * @return whether given format code contains section representing date.
     */
    private static boolean containsDate( String formatCode ) {
        if( containsUnescapedSubstring( formatCode, "y" ) ) {
            return true;
        }
        if( containsUnescapedSubstring( formatCode, "d" ) ) {
            return true;
        }

        formatCode = formatCode.replace( "h:mm", "" );
        formatCode = formatCode.replace( "h:m", "" );
        formatCode = formatCode.replace( "mm:s", "" );
        formatCode = formatCode.replace( "m:s", "" );
        formatCode = formatCode.replace( "am/pm", "" );
        if( containsUnescapedSubstring( formatCode, "m" ) ) {
            return true;
        }
        return false;
    }

    /** Checks whether given format code contains specified substring, which is not escaped using backslash or quotes.
     * @param formatCode format code to analyze. May not be null.
     * @param substring substring to check.
     * @return whether given format code contains specified substring, which is not escaped using backslash or quotes.
     */
    private static boolean containsUnescapedSubstring( String formatCode, String substring ) {
        int index = formatCode.indexOf( substring );
        if( index == 0 ) {
            return true;
        }
        while( index > 0 ) {
            if( isInsideQuotes( formatCode, index ) || isEscapedWithBackslash( formatCode, index ) ) {
                if( index + 1 < formatCode.length() ) {
                    index = formatCode.indexOf( substring, index + 1 );
                    continue;
                } else {
                    break;
                }
            } else {
                return true;
            }
        }
        return false;
    }

    /** Checks whether character at specified index is located inside quotes.
     * @param formatCode format code to analyze. May not be null.
     * @param charIndex index of the character from given format code. Must not be negative.
     * @return whether character at specified index is located inside quotes.
     */
    private static boolean isInsideQuotes( String formatCode, int charIndex ) {
        String before = formatCode.substring( 0, charIndex );
        return countQuoteCharacters( before ) % 2 == 1;
    }

    /** Checks whether character at specified index is escaped with backslash.
     * @param formatCode format code to analyze. May not be null.
     * @param charIndex index of the character from given format code. Must not be negative.
     * @return whether character at specified index is escaped with backslash.
     */
    private static boolean isEscapedWithBackslash( String formatCode, int charIndex ) {
        if( charIndex == 0 ) {
            return false;
        }
        int backslashCount = 0;
        while( charIndex > 0 ) {
            charIndex--;
            if( formatCode.charAt( charIndex ) == '\\' ) {
                backslashCount++;
            } else {
                break;
            }
        }
        return backslashCount % 2 == 1;
    }

    /** Counts quote characters included inside specified format code.
     * @param formatCode format code to analyze. May not be null.
     * @return number of quote characters included inside specified format code.
     */
    private static int countQuoteCharacters( String formatCode ) {
        int count = 0;
        int index = 0;
        while( (index = formatCode.indexOf( "\"", index )) > -1 ) {
            if( !isEscapedWithBackslash( formatCode, index ) ) {
                count++;
            }
            index++;
        }
        return count;
    }
}
