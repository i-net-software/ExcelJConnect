/*
 * Copyright 2024 i-net software
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

/** Exception that may be thrown when I/O or processing errors occur while reading data from Excel document.
 */
public class ExcelParserException extends RuntimeException {

    /** Creates new exception with specified cause.
     * @param cause cause of the exception.
     */
    public ExcelParserException( Throwable cause ) {
        super( cause );
    }

    /** Creates new exception with specified message.
     * @param message the detail message.
     */
    public ExcelParserException( String message ) {
        super( message );
    }
}