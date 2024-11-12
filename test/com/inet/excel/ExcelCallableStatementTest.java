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

import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Paths;

import org.junit.jupiter.api.Test;

import com.inet.excel.parser.ExcelParser;

public class ExcelCallableStatementTest {

    @Test
    public void constructor_throws_exception_if_parser_is_null() {
        assertThrows( IllegalArgumentException.class, () -> new ExcelCallableStatement( null, "Sheet1" ) );
    }

    @Test
    public void constructor_throws_exception_if_sheetName_is_null() {
        ExcelParser parser = new ExcelParser( Paths.get( "zxc.xlsx" ), false );
        assertThrows( IllegalArgumentException.class, () -> new ExcelCallableStatement( parser, null ) );
    }
}
