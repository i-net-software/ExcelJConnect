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

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import com.inet.excel.parser.RowData.CellData;

/** Component responsible for reading data from Excel document.
 */
public class ExcelParser {

    private final XMLInputFactory        factory                         = XMLInputFactory.newInstance();
    private final Path                   filePath;
    private final boolean                hasHeaderRow;

    private List<String>                 sharedStrings                   = null;
    private Map<String, String>          sheetNamesToPaths               = null;
    private List<ValueType>              valueTypesOrderedByStyleIndexes = null;
    private Map<String, SheetDimension>  sheetNamesToDimensions          = new HashMap<>();
    private Map<String, List<String>>    sheetNamesToColumnNames         = new HashMap<>();
    private Map<String, List<ValueType>> sheetNamesToColumnTypes         = new HashMap<>();

    /** Creates instance responsible for reading data from specified Excel document.
     * @param filePath file path to Excel document.
     * @param hasHeaderRow whether first row in sheet represents column headers. 
     * @throws IllegalArgumentException if file path is null.
     */
    public ExcelParser( Path filePath, boolean hasHeaderRow ) {
        if( filePath == null ) {
            throw new IllegalArgumentException( "filePath must not be null" );
        }
        this.filePath = filePath;
        this.hasHeaderRow = hasHeaderRow;
    }

    /** Exception that may be thrown when I/O or processing errors occur while reading data from Excel document.
     */
    public static class ExcelParserException extends RuntimeException {

        /** Creates new exception with specified cause.
         * @param cause cause of the exception.
         */
        public ExcelParserException( Throwable cause ) {
            super( cause );
        }
    }

    /** Returns file name of the Excel document, e.g. "doc.xlsx".
     * @return file name of the Excel document.
     */
    public String getFileName() {
        return filePath.getFileName().toString();
    }

    /** Returns list containing names of columns from specified sheet.
     * If sheet contains row representing column headers, values from its cells will be used as column names.
     * In case of cells in column header row, which have no values, column names will be auto-generated.
     * If sheet does not contain row representing column headers, all column names will be auto-generated.
     * @param sheetName name of the sheet from Excel document.
     * @return list containing names of columns from specified sheet.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public List<String> getColumnNames( String sheetName ) {
        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            initDimensionAndColumnNames( zipFile, sheetName );
            return Collections.unmodifiableList( sheetNamesToColumnNames.get( sheetName ) );
        } catch( IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns list containing names of all sheets from Excel document, in order of their occurrence in the workbook.
     * @return list containing names of all sheets from Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public List<String> getSheetNames() {
        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            return sheetNamesToPaths.entrySet().stream().sorted( Map.Entry.comparingByValue() ).map( Map.Entry::getKey ).collect( Collectors.toList() );
        } catch( IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns list of column types from specified sheet.
     * It probes limited number of cells belonging to columns in order to recognize their common value type.
     * In case of columns with values of mixed types, it takes {@link ValueType#VARCHAR} as column's type.
     * @param sheetName name of the sheet from Excel document.
     * @return list of column types from specified sheet.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public List<ValueType> getColumnTypes( String sheetName ) {
        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            initStyles( zipFile );
            initDimensionAndColumnNames( zipFile, sheetName );
            initColumnTypes( zipFile, sheetName );
            return Collections.unmodifiableList( sheetNamesToColumnTypes.get( sheetName ) );
        } catch( IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns number of rows included in specified sheet from Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @return number of rows included in specified sheet from Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public int getRowCount( String sheetName ) {
        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            int rowCount = readRowCount( zipFile, sheetName );
            if( hasHeaderRow ) {
                // should not count header row
                return Math.max( 0, rowCount - 1 );
            }
            return rowCount;
        } catch( IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns list of rows from specified range. Every element in resulting list represents cell values from single row.
     * Resulting list contains data of rows in order of their occurrence in the sheet. Cells with no values are represented as empty strings.
     * @param sheetName name of the sheet from Excel document.
     * @param firstRowIndex index of the first row, which should be included in the list.
     * @param lastRowIndex index of the last row, which should be included in the list.
     * @return list of rows from specified range.
     * @throws IllegalArgumentException if one of specified indexes is smaller than 1; if first index is greater than last index.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public List<List<Object>> getRows( String sheetName, int firstRowIndex, int lastRowIndex ) {
        if( firstRowIndex < 1 ) {
            throw new IllegalArgumentException( "firstRowIndex must be greater than zero" );
        }
        if( lastRowIndex < 1 ) {
            throw new IllegalArgumentException( "lastRowIndex must be greater than zero" );
        }
        if( firstRowIndex > lastRowIndex ) {
            throw new IllegalArgumentException( "firstRowIndex  must be smaller than or equal to lastRowIndex" );
        }

        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            initStyles( zipFile );
            initDimensionAndColumnNames( zipFile, sheetName );
            if( hasHeaderRow ) {
                // should skip header row
                firstRowIndex++;
                lastRowIndex++;
            }
            return readRows( zipFile, sheetName, firstRowIndex, lastRowIndex );
        } catch( IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Initializes map of sheet names to their paths within Excel document, if these are not already loaded.
     * @param zipFile component allowing access to data inside Excel file.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initSheetData( ZipFile zipFile ) {
        if( sheetNamesToPaths != null ) {
            return;
        }

        try {
            Map<String, String> sheetRelIdToName = new HashMap<>();

            ZipEntry workbookEntry = zipFile.getEntry( "xl/workbook.xml" );
            try( InputStream is = zipFile.getInputStream( workbookEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {
                    while( reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            if( "sheet".equals( reader.getLocalName() ) ) {
                                String rID = reader.getAttributeValue( null, "id" );
                                String sheetName = reader.getAttributeValue( null, "name" );
                                if( rID != null && sheetName != null ) {
                                    sheetRelIdToName.put( rID, sheetName );
                                }
                            }
                        }
                    }
                } finally {
                    reader.close();
                }
            }

            ZipEntry relsEntry = zipFile.getEntry( "xl/_rels/workbook.xml.rels" );
            try( InputStream is = zipFile.getInputStream( relsEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {
                    Map<String, String> map = new HashMap<>();

                    Set<String> sheetRelIds = new HashSet<>( sheetRelIdToName.keySet() );
                    while( !sheetRelIds.isEmpty() && reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            String rID = reader.getAttributeValue( null, "Id" );
                            if( sheetRelIds.contains( rID ) ) {
                                sheetRelIds.remove( rID );

                                String sheetName = sheetRelIdToName.get( rID );
                                String target = reader.getAttributeValue( null, "Target" );
                                if( target == null ) {
                                    String msg = "Relationship of sheet \"" + sheetName + "\" must include attribute \"Target\".";
                                    throw new ExcelParserException( new IllegalStateException( msg ) );
                                }
                                map.put( sheetName, "xl/" + target );
                            }
                        }
                    }

                    sheetNamesToPaths = map;
                } finally {
                    reader.close();
                }
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Initializes list of shared strings, if these are not already loaded.
     * @param zipFile component allowing access to data inside Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initSharedStrings( ZipFile zipFile ) {
        if( sharedStrings != null ) {
            return;
        }

        ZipEntry sheetEntry = zipFile.getEntry( "xl/sharedStrings.xml" );
        try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
            XMLStreamReader reader = factory.createXMLStreamReader( is );
            try {
                List<String> list = new ArrayList<>();

                while( reader.hasNext() ) {
                    reader.next();
                    if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                        String localName = reader.getLocalName();
                        if( "t".equals( localName ) ) {
                            list.add( reader.getElementText() );
                        }
                    }
                }

                sharedStrings = list;
            } finally {
                reader.close();
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Initializes list of value types defined for cells with specific styles, if these are not already loaded.
     * @param zipFile component allowing access to data inside Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initStyles( ZipFile zipFile ) {
        if( valueTypesOrderedByStyleIndexes != null ) {
            return;
        }

        ZipEntry sheetEntry = zipFile.getEntry( "xl/styles.xml" );
        try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
            XMLStreamReader reader = factory.createXMLStreamReader( is );
            try {

                boolean insideNumFmts = false; //NOTE: just in case of some invalid documents ("numFmt" should appear inside "numFmts")
                Map<String, String> numFmtIdToFormatCode = new HashMap<>();

                boolean insideCellXfs = false; //NOTE: important to check because "xf" appears inside "cellStyleXfs" and "cellXfs"
                List<String> numFmtIdsFromCellXfs = new ArrayList<>();

                while( reader.hasNext() ) {
                    reader.next();
                    if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                        String localName = reader.getLocalName();
                        if( localName == null ) {
                            continue;
                        }
                        switch( localName ) {
                            case "numFmts":
                                insideNumFmts = true;
                                break;
                            case "numFmt":
                                if( !insideNumFmts ) {
                                    break;
                                }
                                String numFmtId = reader.getAttributeValue( null, "numFmtId" );
                                if( numFmtId != null ) {
                                    String formatCode = reader.getAttributeValue( null, "formatCode" );
                                    if( formatCode == null ) {
                                        formatCode = "";
                                    }
                                    numFmtIdToFormatCode.put( numFmtId, formatCode );
                                }
                                break;
                            case "cellXfs":
                                insideCellXfs = true;
                                break;
                            case "xf":
                                if( !insideCellXfs ) {
                                    break;
                                }
                                String id = reader.getAttributeValue( null, "numFmtId" );
                                if( id == null ) {
                                    id = "unknown"; //NOTE: in case of invalid document, add dummy value to the list to preserve order of occurrence
                                }
                                numFmtIdsFromCellXfs.add( id );
                                break;
                        }
                    } else if( reader.getEventType() == XMLStreamReader.END_ELEMENT ) {
                        String localName = reader.getLocalName();
                        switch( localName ) {
                            case "numFmts":
                                insideNumFmts = false;
                                break;
                            case "cellXfs":
                                insideCellXfs = false;
                                break;
                        }
                    }
                }
                valueTypesOrderedByStyleIndexes = new ArrayList<>();

                for( int styleIndex = 0; styleIndex < numFmtIdsFromCellXfs.size(); styleIndex++ ) {
                    String id = numFmtIdsFromCellXfs.get( styleIndex );
                    try {
                        int intID = Integer.parseInt( id );
                        switch( intID ) {
                            case 14:
                            case 22:
                                valueTypesOrderedByStyleIndexes.add( ValueType.TIMESTAMP );
                                continue;
                            case 15:
                            case 16:
                            case 17:
                                valueTypesOrderedByStyleIndexes.add( ValueType.DATE );
                                continue;
                            case 18:
                            case 19:
                            case 20:
                            case 21:
                                valueTypesOrderedByStyleIndexes.add( ValueType.TIME );
                                continue;
                        }
                    } catch( NumberFormatException ex ) {
                        // ignore
                    }
                    String formatCode = numFmtIdToFormatCode.getOrDefault( id, "" );
                    valueTypesOrderedByStyleIndexes.add( FormatCodeAnalyzer.recognizeValueType( formatCode ) );
                }
            } finally {
                reader.close();
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns zip file entry for specified sheet or throws exception if it is null or such sheet does not exist inside Excel document.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @return zip file entry for specified sheet.
     * @throws ExcelParserException if specified sheet is null or does not exist inside Excel document.
     */
    private ZipEntry getZipEntryForSheet( ZipFile zipFile, String sheetName ) {
        if( sheetName == null ) {
            throw new ExcelParserException( new IllegalArgumentException( "Sheet name must not be null." ) );
        }
        ZipEntry sheetEntry = null;
        String sheetPath = sheetNamesToPaths.get( sheetName );
        if( sheetPath != null ) {
            sheetEntry = zipFile.getEntry( sheetPath );
        }
        if( sheetEntry == null ) {
            String msg = "There is no sheet with name \"" + sheetName + "\".";
            throw new ExcelParserException( new IllegalArgumentException( msg ) );
        }
        return sheetEntry;
    }

    /** Initializes dimension and list of column names from specified sheet.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initDimensionAndColumnNames( ZipFile zipFile, String sheetName ) {
        if( sheetNamesToDimensions.get( sheetName ) != null && sheetNamesToColumnNames.get( sheetName ) != null ) {
            return;
        }
        try {
            ZipEntry sheetEntry = getZipEntryForSheet( zipFile, sheetName );
            try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {

                    boolean insideHeaderRow = false;
                    boolean collectCellRefs = false;

                    RowData headerData = new RowData( 1 );
                    CellData currentCellData = null;
                    SheetDimension sheetDimension = null;
                    RowSpanData rowSpan = new RowSpanData();

                    while( reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( localName == null ) {
                                continue;
                            }
                            switch( localName ) {
                                case "dimension":
                                    String ref = reader.getAttributeValue( null, "ref" );
                                    sheetDimension = SheetDimension.parse( ref );
                                    break;
                                case "row":
                                    collectCellRefs = false;

                                    if( hasHeaderRow ) {
                                        String rowIndex = reader.getAttributeValue( null, "r" );
                                        if( "1".equals( rowIndex ) ) {
                                            insideHeaderRow = true;
                                        }
                                    }

                                    if( sheetDimension == null ) {
                                        String spans = reader.getAttributeValue( null, "spans" );
                                        if( spans == null ) {
                                            collectCellRefs = true;
                                        } else {
                                            rowSpan.addSpanRange( spans );
                                        }
                                    }
                                    break;
                                case "c":
                                    if( insideHeaderRow ) {
                                        currentCellData = new CellData();
                                        currentCellData.setR( reader.getAttributeValue( null, "r" ) );
                                        currentCellData.setT( reader.getAttributeValue( null, "t" ) );
                                        currentCellData.setS( reader.getAttributeValue( null, "s" ) );
                                    }
                                    if( collectCellRefs ) {
                                        String cellRef = reader.getAttributeValue( null, "r" );
                                        rowSpan.addCellRef( cellRef );
                                    }
                                    break;
                                case "v":
                                    if( insideHeaderRow ) {
                                        currentCellData.setV( reader.getElementText() );
                                        headerData.addCellData( currentCellData );
                                        currentCellData = null;
                                    }
                                    break;
                            }
                        } else if( reader.getEventType() == XMLStreamReader.END_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( insideHeaderRow && "row".equals( localName ) ) {
                                insideHeaderRow = false;
                            }
                        }

                        if( !hasHeaderRow && sheetDimension != null ) {
                            break; // we have already all data required to generate column names 
                        }
                    }

                    if( sheetDimension == null ) {
                        if( rowSpan.isEmpty() ) {
                            // sheet is empty
                            sheetNamesToDimensions.put( sheetName, new SheetDimension( 1, 1 ) );
                            sheetNamesToColumnNames.put( sheetName, Collections.singletonList( "C1" ) );
                            return;
                        } else {
                            sheetDimension = new SheetDimension( rowSpan.getFirstColumnIndex(), rowSpan.getLastColumnIndex() );
                        }
                    }
                    List<String> columnNames = generateColumnNames( sheetDimension.getFirstColumnIndex(), sheetDimension.getLastColumnIndex() );

                    for( CellData cell : headerData.getCellsInRow() ) { // NOTE: relevant only if hasHeaderRow is true
                        Object value = getCellValue( zipFile, cell );
                        if( value == null ) {
                            continue;
                        }
                        int columnIndex = SheetDimension.getColumnIndexFromCellRef( cell.getR() );
                        if( columnIndex > 0 ) { // ensures that cell ref is valid
                            columnIndex -= sheetDimension.getFirstColumnIndex();
                            if( columnIndex >= 0 && columnIndex < columnNames.size() ) {
                                columnNames.set( columnIndex, value.toString() );
                            }
                        }
                    }
                    sheetNamesToDimensions.put( sheetName, sheetDimension );
                    sheetNamesToColumnNames.put( sheetName, columnNames );
                    return;
                } finally {
                    reader.close();
                }
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns list containing auto-generated column names for specified range.
     * @param firstColumnIndex index of the first column, which name should be included in the list.
     * @param lastColumnIndex index of the last column, which name should be included in the list.
     * @return list containing auto-generated column names for specified range.
     */
    private List<String> generateColumnNames( int firstColumnIndex, int lastColumnIndex ) {
        List<String> columnNames = new ArrayList<>();
        for( int index = firstColumnIndex; index <= lastColumnIndex; index++ ) {
            columnNames.add( "C" + index );
        }
        return columnNames;
    }

    /** Initializes list of column types from specified sheet.
     * It probes limited number of cells belonging to columns in order to recognize their common value type.
     * In case of columns with values of mixed types, it takes {@link ValueType#VARCHAR} as column's type.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initColumnTypes( ZipFile zipFile, String sheetName ) {
        if( sheetNamesToColumnTypes.get( sheetName ) != null ) {
            return;
        }
        try {
            ZipEntry sheetEntry = getZipEntryForSheet( zipFile, sheetName );
            try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {
                    int columnCount = sheetNamesToColumnNames.get( sheetName ).size();
                    SheetDimension sheetDimension = sheetNamesToDimensions.get( sheetName );

                    final int probedCellLimit = 10;
                    final int probedRowLimit = 30;

                    int[] probedCells = new int[columnCount];
                    ValueType[] valueTypes = new ValueType[columnCount];
                    int probedRowCount = 0;
                    boolean insideRow = false;

                    ValueType typeToSet = null;
                    Optional<Integer> columnIndexOfValueToCheck = Optional.empty(); 
                    boolean possibleNumber = false;

                    while( reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( localName == null ) {
                                continue;
                            }
                            switch( localName ) {
                                case "row":
                                    if( hasHeaderRow ) {
                                        String rowIndex = reader.getAttributeValue( null, "r" );
                                        if( "1".equals( rowIndex ) ) {
                                            break; // skip header row
                                        }
                                    }
                                    insideRow = true;
                                    break;
                                case "c":
                                    if( !insideRow ) {
                                        break;
                                    }
                                    columnIndexOfValueToCheck = Optional.empty();
                                    possibleNumber = false;
                                    typeToSet = null;

                                    String cellRef = reader.getAttributeValue( null, "r" );
                                    int columnIndex = SheetDimension.getColumnIndexFromCellRef( cellRef );
                                    if( columnIndex > 0 ) { // ensures that cell ref is valid
                                        columnIndex -= sheetDimension.getFirstColumnIndex();
                                        if( columnIndex >= 0 && columnIndex < columnCount ) {
                                            if( probedCells[columnIndex] == probedCellLimit ) {
                                                break; // probed enough cells
                                            }
                                            if( valueTypes[columnIndex] == ValueType.VARCHAR ) {
                                                break; // already initialized as most general type
                                            }

                                            columnIndexOfValueToCheck = Optional.of( Integer.valueOf( columnIndex ) );

                                            if( "s".equals( reader.getAttributeValue( null, "t" ) ) ) {
                                                typeToSet = ValueType.VARCHAR;
                                            } else {
                                                try {
                                                    int styleIndex = Integer.parseInt( reader.getAttributeValue( null, "s" ) );
                                                    typeToSet = valueTypesOrderedByStyleIndexes.get( styleIndex );
                                                } catch( NumberFormatException | IndexOutOfBoundsException ex ) {
                                                    // since style could not be recognized, VARCHAR stays as column's type
                                                    typeToSet = ValueType.VARCHAR;
                                                }
                                                ValueType currentType = valueTypes[columnIndex];
                                                if( ( currentType == null || currentType == ValueType.NUMBER ) && typeToSet == ValueType.VARCHAR ) {
                                                    // type of cell value is VARCHAR, but it will check whether value can be parsed as number
                                                    possibleNumber = true;
                                                } else if( currentType != null && currentType != typeToSet ) {
                                                    // column has values of various types so it will set VARCHAR as column type
                                                    typeToSet = ValueType.VARCHAR;
                                                }
                                            }
                                        }
                                    }
                                    break;
                                case "v":
                                    if( columnIndexOfValueToCheck.isPresent() ) {
                                        String value = reader.getElementText();
                                        if( value == null || value.trim().isEmpty() ) {
                                            break;
                                        }
                                        int colIndex = columnIndexOfValueToCheck.get().intValue();

                                        if( possibleNumber ) {
                                            try {
                                                Double.parseDouble( value );
                                                typeToSet = ValueType.NUMBER;
                                            } catch( NullPointerException | NumberFormatException ex ) {
                                                typeToSet = ValueType.VARCHAR;
                                            }
                                            possibleNumber = false;
                                        } 
                                        valueTypes[colIndex] = typeToSet;
                                        probedCells[colIndex]++;
                                        columnIndexOfValueToCheck = Optional.empty();
                                    }
                                    break;
                                default:
                                    break;
                            }
                        } else if( reader.getEventType() == XMLStreamReader.END_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( insideRow && "row".equals( localName ) ) {
                                insideRow = false;
                                probedRowCount++;
                                if( probedRowCount == probedRowLimit ) {
                                    break; // probed enough rows
                                }
                            }
                        }
                    }

                    for( int index = 0; index < valueTypes.length; index++ ) {
                        if( valueTypes[index] == null ) {
                            valueTypes[index] = ValueType.VARCHAR; // fallback to string
                        }
                    }
                    sheetNamesToColumnTypes.put( sheetName, Arrays.asList( valueTypes ) );
                } finally {
                    reader.close();
                }
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns list of rows from specified range. Every element in resulting list represents cell values from single row.
     * Resulting list contains data of rows in order of their occurrence in the sheet. Cells with no values are represented as empty strings.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @param firstRowIndex index of the first row, which should be included in the list.
     * @param lastRowIndex index of the last row, which should be included in the list.
     * @return list of rows from specified range.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private List<List<Object>> readRows( ZipFile zipFile, String sheetName, int firstRowIndex, int lastRowIndex ) {
        try {
            ZipEntry sheetEntry = getZipEntryForSheet( zipFile, sheetName );
            try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {
                    int requestedRowCount = lastRowIndex - firstRowIndex + 1;
                    int columnCount = sheetNamesToColumnNames.get( sheetName ).size();
                    SheetDimension sheetDimension = sheetNamesToDimensions.get( sheetName );

                    List<List<Object>> allRows = new ArrayList<>();
                    IntStream.range( 0, requestedRowCount ).forEach( val -> {
                        List<Object> row = new ArrayList<>();
                        IntStream.range( 0, columnCount ).forEach( v -> row.add( null ) );
                        allRows.add( row );
                    } );

                    RowData currentRowData = null;
                    CellData currentCellData = null;

                    while( reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( localName == null ) {
                                continue;
                            }
                            switch( localName ) {
                                case "row":
                                    try {
                                        int rowIndex = Integer.parseInt( reader.getAttributeValue( null, "r" ) );
                                        if( rowIndex >= firstRowIndex && rowIndex <= lastRowIndex ) {
                                            currentRowData = new RowData( rowIndex );
                                        }
                                    } catch( Exception ex ) {
                                        // ignore row if index can not be parsed
                                    }
                                    break;
                                case "c":
                                    if( currentRowData != null ) {
                                        currentCellData = new CellData();
                                        currentCellData.setR( reader.getAttributeValue( null, "r" ) );
                                        currentCellData.setT( reader.getAttributeValue( null, "t" ) );
                                        currentCellData.setS( reader.getAttributeValue( null, "s" ) );
                                    }
                                    break;
                                case "v":
                                    if( currentRowData != null ) {
                                        currentCellData.setV( reader.getElementText() );
                                        currentRowData.addCellData( currentCellData );
                                        currentCellData = null;
                                    }
                                    break;
                            }
                        } else if( reader.getEventType() == XMLStreamReader.END_ELEMENT ) {
                            String localName = reader.getLocalName();
                            if( "row".equals( localName ) ) {
                                if( currentRowData != null ) {
                                    List<Object> row = allRows.get( currentRowData.getRowIndex() - firstRowIndex );

                                    for( CellData cell : currentRowData.getCellsInRow() ) {
                                        Object value = getCellValue( zipFile, cell );
                                        if( value == null ) {
                                            continue;
                                        }

                                        int columnIndex = SheetDimension.getColumnIndexFromCellRef( cell.getR() );
                                        if( columnIndex > 0 ) { // ensures that cell ref is valid
                                            columnIndex -= sheetDimension.getFirstColumnIndex();
                                            if( columnIndex >= 0 && columnIndex < columnCount ) {
                                                row.set( columnIndex, value );
                                            }
                                        }
                                    }
                                    currentRowData = null;
                                }
                            }
                        }
                    }
                    return allRows;
                } finally {
                    reader.close();
                }
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Returns value of specified cell.
     * @param zipFile component allowing access to data inside Excel document.
     * @param cell container with data of the cell.
     * @return value of specified cell or null, in case of invalid data.
     */
    private Object getCellValue( ZipFile zipFile, CellData cell ) {
        if( "s".equals( cell.getT() ) ) {
            try {
                int index = Integer.parseInt( cell.getV() );
                initSharedStrings( zipFile );
                return sharedStrings.get( index );
            } catch( NumberFormatException ex ) {
                return null;
            }
        } else {
            int styleIndex;
            try {
                styleIndex = Integer.parseInt( cell.getS() );
            } catch( Exception ex ) {
                return cell.getV(); // fallback to string
            }

            Function<CellData, Long> toMillis = cellData -> {
                double value = Double.parseDouble( cellData.getV() );
                int days = Double.valueOf( value ).intValue();
                int seconds = Long.valueOf( Math.round( (value - days) * TimeUnit.DAYS.toSeconds( 1 ) ) ).intValue();

                Calendar cal = Calendar.getInstance();
                days--; // because value 0 represents "0 January 1900" in excel
                cal.set( 1900, 0, days, 0, 0, seconds );
                cal.set( Calendar.MILLISECOND, 0 );
                return new Long( cal.getTime().getTime() );
            };

            ValueType valueType = valueTypesOrderedByStyleIndexes.get( styleIndex );
            switch( valueType ) {
                case DATE:
                    try {
                        return new Date( toMillis.apply( cell ).longValue() );
                    } catch( Exception ex ) {
                        return null;
                    }
                case TIME:
                    try {
                        return new Time( toMillis.apply( cell ).longValue() );
                    } catch( Exception ex ) {
                        return null;
                    }
                case TIMESTAMP:
                    try {
                        return new Timestamp( toMillis.apply( cell ).longValue() );
                    } catch( Exception ex ) {
                        return null;
                    }
                case NUMBER:
                    try {
                        return Double.valueOf( cell.getV() );
                    } catch( Exception ex ) {
                        return null;
                    }
                case VARCHAR:
                default:
                    return cell.getV();
            }
        }
    }

    /** Returns number of rows included in specified sheet from Excel document.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @return number of rows included in specified sheet from Excel document.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private int readRowCount( ZipFile zipFile, String sheetName ) {
        try {
            ZipEntry sheetEntry = getZipEntryForSheet( zipFile, sheetName );
            try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {
                    int rowCount = 0;
                    while( reader.hasNext() ) {
                        reader.next();
                        if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                            if( "row".equals( reader.getLocalName() ) ) {
                                try {
                                    int rowIndex = Integer.parseInt( reader.getAttributeValue( null, "r" ) );
                                    if( rowIndex > rowCount ) {
                                        rowCount = rowIndex;
                                    }
                                } catch( Exception ex ) {
                                    // ignore row if index can not be parsed
                                }
                            }
                        }
                    }
                    return rowCount;
                } finally {
                    reader.close();
                }
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }
}
