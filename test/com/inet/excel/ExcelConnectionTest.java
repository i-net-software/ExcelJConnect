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

import java.nio.file.Paths;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser;

public class ExcelConnectionTest {

    @Test
    public void constructor_throws_exception_if_parser_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelConnection( null, null ) );
    }

    @Test
    public void constructor_accepts_null_as_runnable_to_be_executed_on_connection_close() {
        ExcelParser parser = new ExcelParser( Paths.get( "" ), false );
        assertDoesNotThrow( () -> new ExcelConnection( parser, null ) );
    }
}
