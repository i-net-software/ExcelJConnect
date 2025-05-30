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

import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.RowIdLifetime;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

import com.inet.excel.parser.ExcelParser;
import com.inet.excel.parser.ValueType;

/** Implementation of {@link ResultSetMetaData} for {@link ExcelConnection}
 */
public class ExcelDatabaseMetaData implements DatabaseMetaData {

    public static final int COLUMN_SIZE_IN_BYTES = 65535;

    public static final String PRODUCT_NAME = "inetexcel";

    private final ExcelParser parser;

    /** Constructor of the class.
     * @param parser component responsible for reading data from Excel document.
     * @throws IllegalArgumentException if given parser is null.
     */
    public ExcelDatabaseMetaData( ExcelParser parser ) {
        if( parser == null ) {
            throw new IllegalArgumentException( "parser must not be null" );
        }
        this.parser = parser;
    }

    /** Returns constant value that identifies the generic SQL type matching specified {@link ValueType}.
     * @param type type to return constant value for. Must not be null.
     * @return constant value that identifies the generic SQL type matching specified {@link ValueType}.
     */
    static int getDataType( ValueType type ) {
        switch( type ) {
            case DATE:
                return Types.DATE;
            case NUMBER:
                return Types.NUMERIC;
            case TIME:
                return Types.TIME;
            case TIMESTAMP:
                return Types.TIMESTAMP;
            case VARCHAR:
                return Types.VARCHAR;
            default:
                return Types.JAVA_OBJECT;
        }
    }

    /** Returns name that identifies the generic SQL type matching specified {@link ValueType}.
     * @param type type to return constant value for. Must not be null.
     * @return name that identifies the generic SQL type matching specified {@link ValueType}.
     */
    static String getDataTypeName( ValueType type ) {
        switch( type ) {
            case DATE:
                return "DATE";
            case NUMBER:
                return "NUMERIC";
            case TIME:
                return "TIME";
            case TIMESTAMP:
                return "TIMESTAMP";
            case VARCHAR:
                return "VARCHAR";
            default:
                return "JAVA_OBJECT";
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getTables( String catalog, String schemaPattern, String tableNamePattern, String[] types ) throws SQLException {
        List<String> columnNames = Arrays.asList( "TABLE_CAT", "TABLE_SCHEM", "TABLE_NAME", "TABLE_TYPE", "REMARKS", "TYPE_CAT", "TYPE_SCHEM", "TYPE_NAME", "SELF_REFERENCING_COL_NAME", "REF_GENERATION" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getColumns( String catalog, String schemaPattern, String tableNamePattern, String columnNamePattern ) throws SQLException {
        List<String> columnNames = Arrays.asList( "TABLE_CAT", "TABLE_SCHEM", "TABLE_NAME", "COLUMN_NAME", "DATA_TYPE", "TYPE_NAME", "COLUMN_SIZE", "BUFFER_LENGTH", //
                                                  "DECIMAL_DIGITS", "NUM_PREC_RADIX", "NULLABLE", "REMARKS", "COLUMN_DEF", "SQL_DATA_TYPE", "SQL_DATETIME_SUB", //
                                                  "CHAR_OCTET_LENGTH", "ORDINAL_POSITION", "IS_NULLABLE", "SCOPE_CATALOG", "SCOPE_SCHEMA", "SCOPE_TABLE", //
                                                  "SOURCE_DATA_TYPE", "IS_AUTOINCREMENT", "IS_GENERATEDCOLUMN" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getDatabaseProductName() throws SQLException {
        return PRODUCT_NAME;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public <T> T unwrap( Class<T> iface ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isWrapperFor( Class<?> iface ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean allProceduresAreCallable() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean allTablesAreSelectable() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getURL() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getUserName() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isReadOnly() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean nullsAreSortedHigh() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean nullsAreSortedLow() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean nullsAreSortedAtStart() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean nullsAreSortedAtEnd() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getDatabaseProductVersion() throws SQLException {
        return ExcelDriver.MAJOR_VERSION + "." + ExcelDriver.MINOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getDriverName() throws SQLException {
        return ExcelDriver.DRIVER_NAME;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getDriverVersion() throws SQLException {
        return getDriverMajorVersion() + "." + getDriverMinorVersion();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getDriverMajorVersion() {
        return ExcelDriver.MAJOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getDriverMinorVersion() {
        return ExcelDriver.MINOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean usesLocalFiles() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean usesLocalFilePerTable() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMixedCaseIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesUpperCaseIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesLowerCaseIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesMixedCaseIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMixedCaseQuotedIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesUpperCaseQuotedIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesLowerCaseQuotedIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean storesMixedCaseQuotedIdentifiers() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getIdentifierQuoteString() throws SQLException {
        return " ";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSQLKeywords() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getNumericFunctions() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getStringFunctions() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSystemFunctions() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getTimeDateFunctions() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSearchStringEscape() throws SQLException {
        return "";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getExtraNameCharacters() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsAlterTableWithAddColumn() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsAlterTableWithDropColumn() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsColumnAliasing() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean nullPlusNonNullIsNull() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsConvert() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsConvert( int fromType, int toType ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsTableCorrelationNames() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsDifferentTableCorrelationNames() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsExpressionsInOrderBy() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOrderByUnrelated() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsGroupBy() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsGroupByUnrelated() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsGroupByBeyondSelect() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsLikeEscapeClause() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMultipleResultSets() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMultipleTransactions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsNonNullableColumns() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMinimumSQLGrammar() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCoreSQLGrammar() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsExtendedSQLGrammar() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsANSI92EntryLevelSQL() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsANSI92IntermediateSQL() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsANSI92FullSQL() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsIntegrityEnhancementFacility() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOuterJoins() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsFullOuterJoins() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsLimitedOuterJoins() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSchemaTerm() throws SQLException {
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getProcedureTerm() throws SQLException {
        return "sheet";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getCatalogTerm() throws SQLException {
        return "file";
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isCatalogAtStart() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getCatalogSeparator() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSchemasInDataManipulation() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSchemasInProcedureCalls() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSchemasInTableDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSchemasInIndexDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSchemasInPrivilegeDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCatalogsInDataManipulation() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCatalogsInProcedureCalls() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCatalogsInTableDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCatalogsInIndexDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCatalogsInPrivilegeDefinitions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsPositionedDelete() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsPositionedUpdate() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSelectForUpdate() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsStoredProcedures() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSubqueriesInComparisons() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSubqueriesInExists() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSubqueriesInIns() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSubqueriesInQuantifieds() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsCorrelatedSubqueries() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsUnion() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsUnionAll() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOpenCursorsAcrossCommit() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOpenCursorsAcrossRollback() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOpenStatementsAcrossCommit() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsOpenStatementsAcrossRollback() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxBinaryLiteralLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxCharLiteralLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnsInGroupBy() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnsInIndex() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnsInOrderBy() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnsInSelect() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxColumnsInTable() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxConnections() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxCursorNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxIndexLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxSchemaNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxProcedureNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxCatalogNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxRowSize() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean doesMaxRowSizeIncludeBlobs() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxStatementLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxStatements() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxTableNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxTablesInSelect() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getMaxUserNameLength() throws SQLException {
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getDefaultTransactionIsolation() throws SQLException {
        return Connection.TRANSACTION_NONE;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsTransactions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsTransactionIsolationLevel( int level ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsDataDefinitionAndDataManipulationTransactions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsDataManipulationTransactionsOnly() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean dataDefinitionCausesTransactionCommit() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean dataDefinitionIgnoredInTransactions() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getProcedures( String catalog, String schemaPattern, String procedureNamePattern ) throws SQLException {
        List<String> columnNames = Arrays.asList( "PROCEDURE_CAT", "PROCEDURE_SCHEM", "PROCEDURE_NAME", //
                                                  "reserved for future use 1", "reserved for future use 2", "reserved for future use 3",
                                                  "REMARKS", "PROCEDURE_TYPE", "SPECIFIC_NAME" );
        List<List<Object>> allRows = new ArrayList<>();

        try {
            for( String sheetName : parser.getSheetNames() ) {
                List<Object> row = new ArrayList<>();
                row.add( parser.getFileName() );
                row.add( null );
                row.add( sheetName );
                row.add( null );
                row.add( null );
                row.add( null );
                row.add( "" );
                row.add( Integer.valueOf( DatabaseMetaData.procedureReturnsResult ) );
                row.add( null );

                allRows.add( row );
            }
        } catch( Exception ex ) {
            throw new SQLException( ex );
        }

        return new ExcelDatabaseResultSet( columnNames, allRows );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getProcedureColumns( String catalog, String schemaPattern, String procedureNamePattern, String columnNamePattern ) throws SQLException {
        List<String> columnNames = Arrays.asList( "PROCEDURE_CAT", "PROCEDURE_SCHEM", "PROCEDURE_NAME", "COLUMN_NAME", "COLUMN_TYPE", //
                                                  "DATA_TYPE", "TYPE_NAME", "PRECISION", "LENGTH", "SCALE", "RADIX", "NULLABLE", //
                                                  "REMARKS", "COLUMN_DEF", "SQL_DATA_TYPE", "SQL_DATETIME_SUB", "CHAR_OCTET_LENGTH", //
                                                  "ORDINAL_POSITION", "IS_NULLABLE", "SPECIFIC_NAME" );
        List<List<Object>> allRows = new ArrayList<>();

        try {
            for( String sheetName : parser.getSheetNames() ) {
                if( procedureNamePattern != null && !Objects.equals( sheetName, procedureNamePattern ) ) {
                    continue;
                }

                List<ValueType> columnTypes = parser.getColumnTypes( sheetName );
                int colIndex = 0;
                for( String colName : parser.getColumnNames( sheetName ) ) {
                    ValueType valueType = columnTypes.get( colIndex );

                    colIndex++; // indexing starts with 1

                    List<Object> row = new ArrayList<>();
                    row.add( parser.getFileName() );
                    row.add( null );
                    row.add( sheetName );
                    row.add( colName );
                    row.add( Integer.valueOf( DatabaseMetaData.procedureColumnResult ) );
                    row.add( Integer.valueOf( getDataType( valueType ) ) );
                    row.add( getDataTypeName( valueType ) );
                    row.add( Integer.valueOf( 0 ) );
                    row.add( Integer.valueOf( COLUMN_SIZE_IN_BYTES ) );
                    row.add( Integer.valueOf( COLUMN_SIZE_IN_BYTES ) );
                    row.add( null );
                    row.add( Integer.valueOf( DatabaseMetaData.procedureNullable ) );
                    row.add( "" );
                    row.add( "" );
                    row.add( null );
                    row.add( null );
                    row.add( Integer.valueOf( COLUMN_SIZE_IN_BYTES ) );
                    row.add( Integer.valueOf( colIndex ) );
                    row.add( "YES" );
                    row.add( sheetName );

                    allRows.add( row );
                }
            }
        } catch( Exception ex ) {
            throw new SQLException( ex );
        }

        return new ExcelDatabaseResultSet( columnNames, allRows );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getSchemas() throws SQLException {
        List<String> columnNames = Arrays.asList( "TABLE_SCHEM", "TABLE_CATALOG" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getCatalogs() throws SQLException {
        List<String> columnNames = Arrays.asList( "TABLE_CAT" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getTableTypes() throws SQLException {
        List<String> columnNames = Arrays.asList( "TABLE_TYPE" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getColumnPrivileges( String catalog, String schema, String table, String columnNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getTablePrivileges( String catalog, String schemaPattern, String tableNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getBestRowIdentifier( String catalog, String schema, String table, int scope, boolean nullable ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getVersionColumns( String catalog, String schema, String table ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getPrimaryKeys( String catalog, String schema, String table ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getImportedKeys( String catalog, String schema, String table ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getExportedKeys( String catalog, String schema, String table ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getCrossReference( String parentCatalog, String parentSchema, String parentTable, String foreignCatalog, String foreignSchema, String foreignTable ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getTypeInfo() throws SQLException {
        List<String> columnNames = Arrays.asList( "TYPE_NAME", "DATA_TYPE", "PRECISION", "LITERAL_PREFIX", "LITERAL_SUFFIX", "CREATE_PARAMS", "NULLABLE", "CASE_SENSITIVE", "SEARCHABLE", //
                                                  "UNSIGNED_ATTRIBUTE", "FIXED_PREC_SCALE", "AUTO_INCREMENT", "LOCAL_TYPE_NAME", "MINIMUM_SCALE", "MAXIMUM_SCALE", //
                                                  "SQL_DATA_TYPE", "SQL_DATETIME_SUB", "NUM_PREC_RADIX" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getIndexInfo( String catalog, String schema, String table, boolean unique, boolean approximate ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsResultSetType( int type ) throws SQLException {
        return type == ResultSet.TYPE_FORWARD_ONLY;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsResultSetConcurrency( int type, int concurrency ) throws SQLException {
        return type == ResultSet.TYPE_FORWARD_ONLY && concurrency == ResultSet.CONCUR_READ_ONLY;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean ownUpdatesAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean ownDeletesAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean ownInsertsAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean othersUpdatesAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean othersDeletesAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean othersInsertsAreVisible( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean updatesAreDetected( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean deletesAreDetected( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean insertsAreDetected( int type ) throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsBatchUpdates() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getUDTs( String catalog, String schemaPattern, String typeNamePattern, int[] types ) throws SQLException {
        List<String> columnNames = Arrays.asList( "TYPE_CAT", "TYPE_SCHEM", "TYPE_NAME", "CLASS_NAME", "DATA_TYPE", "REMARKS", "BASE_TYPE" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Connection getConnection() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsSavepoints() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsNamedParameters() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsMultipleOpenResults() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsGetGeneratedKeys() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getSuperTypes( String catalog, String schemaPattern, String typeNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getSuperTables( String catalog, String schemaPattern, String tableNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getAttributes( String catalog, String schemaPattern, String typeNamePattern, String attributeNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsResultSetHoldability( int holdability ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getResultSetHoldability() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getDatabaseMajorVersion() throws SQLException {
        return ExcelDriver.MAJOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getDatabaseMinorVersion() throws SQLException {
        return ExcelDriver.MINOR_VERSION;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getJDBCMajorVersion() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getJDBCMinorVersion() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getSQLStateType() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return 0;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean locatorsUpdateCopy() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsStatementPooling() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RowIdLifetime getRowIdLifetime() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getSchemas( String catalog, String schemaPattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean supportsStoredFunctionsUsingCallSyntax() throws SQLException {
        return true;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean autoCommitFailureClosesAllResultSets() throws SQLException {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getClientInfoProperties() throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getFunctions( String catalog, String schemaPattern, String functionNamePattern ) throws SQLException {
        List<String> columnNames = Arrays.asList( "FUNCTION_CAT", "FUNCTION_SCHEM", "FUNCTION_NAME", "REMARKS", "FUNCTION_TYPE", "SPECIFIC_NAME" );
        return new ExcelDatabaseResultSet( columnNames, new ArrayList<>() );
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getFunctionColumns( String catalog, String schemaPattern, String functionNamePattern, String columnNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public ResultSet getPseudoColumns( String catalog, String schemaPattern, String tableNamePattern, String columnNamePattern ) throws SQLException {
        ExcelDriver.throwExceptionAboutUnsupportedOperation();
        return null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean generatedKeyAlwaysReturned() throws SQLException {
        return false;
    }

}
