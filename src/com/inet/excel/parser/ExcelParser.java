package com.inet.excel.parser;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import com.inet.excel.parser.RowData.CellData;

/** Component responsible for reading data from Excel document.
 */
public class ExcelParser {

    private final XMLInputFactory factory = XMLInputFactory.newInstance();
    private final Path filePath;

    private List<String> sharedStrings = null;
    private Map<String, String> sheetNamesToPaths = null;

    /** Creates instance responsible for reading data from specified Excel document.
     * @param filePath file path to Excel document.
     * @throws IllegalArgumentException if file path is null.
     */
    public ExcelParser( Path filePath ) {
        if( filePath == null ) {
            throw new IllegalArgumentException( "filePath must not be null" );
        }
        this.filePath = filePath;
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
     * @param hasHeaderRow whether first row of specified sheet represents column headers. 
     * @return list containing names of columns from specified sheet.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    public List<String> getColumnNames( String sheetName, boolean hasHeaderRow ) {
        try( ZipFile zipFile = new ZipFile( filePath.toString() ) ) {
            initSheetData( zipFile );
            return readColumnNames( zipFile, sheetName, hasHeaderRow );
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

    /** Initializes map of sheet names to their paths within Excel document, if these are not already loaded.
     * @param zipFile component allowing access to data inside Excel file.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private void initSheetData( ZipFile zipFile ) {
        if( sheetNamesToPaths != null ) {
            return;
        }
        sheetNamesToPaths = new HashMap<>();

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
                                sheetNamesToPaths.put( sheetName, "xl/" + target );
                            }
                        }
                    }
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
        sharedStrings = new ArrayList<>();

        ZipEntry sheetEntry = zipFile.getEntry( "xl/sharedStrings.xml" );
        try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
            XMLStreamReader reader = factory.createXMLStreamReader( is );
            try {
                while( reader.hasNext() ) {
                    reader.next();
                    if( reader.getEventType() == XMLStreamReader.START_ELEMENT ) {
                        String localName = reader.getLocalName();
                        if( "t".equals( localName ) ) {
                            sharedStrings.add( reader.getElementText() );
                        }
                    }
                }
            } finally {
                reader.close();
            }
        } catch( XMLStreamException | IOException ex ) {
            throw new ExcelParserException( ex );
        }
    }

    /** Reads names of columns from specified sheet.
     * @param zipFile component allowing access to data inside Excel document.
     * @param sheetName name of the sheet from Excel document.
     * @param hasHeaderRow whether first row of specified sheet represents column headers. 
     * @return list containing names of columns from specified sheet.
     * @throws ExcelParserException in case of I/O or processing errors.
     */
    private List<String> readColumnNames( ZipFile zipFile, String sheetName, boolean hasHeaderRow ) {
        try {
            ZipEntry sheetEntry = null;
            String sheetPath = sheetNamesToPaths.get( sheetName );
            if( sheetPath != null ) {
                sheetEntry = zipFile.getEntry( sheetPath );
            }
            if( sheetEntry == null ) {
                String msg = "There is no sheet with name \"" + sheetName + "\".";
                throw new ExcelParserException( new IllegalArgumentException( msg ) ); //TODO rethink exception type
            }

            try( InputStream is = zipFile.getInputStream( sheetEntry ) ) {
                XMLStreamReader reader = factory.createXMLStreamReader( is );
                try {

                    boolean insideHeaderRow = false;
                    boolean collectCellRefs = false;

                    RowData headerData = new RowData();
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
                            return Collections.singletonList( "C1" ); // sheet is empty
                        } else {
                            sheetDimension = new SheetDimension( rowSpan.getFirstColumnIndex(), rowSpan.getLastColumnIndex() );
                        }
                    }
                    List<String> columnNames = generateColumnNames( sheetDimension.getFirstColumnIndex(), sheetDimension.getLastColumnIndex() );

                    for( CellData cell : headerData.getCellsInRow() ) { // NOTE: relevant only if hasHeaderRow is true
                        String value = null;

                        if( "s".equals( cell.getT() ) ) {
                            try {
                                int index = Integer.parseInt( cell.getV() );
                                initSharedStrings( zipFile );
                                value = sharedStrings.get( index );
                            } catch( NumberFormatException ex ) {
                                // error can be ignored -> column is going to get auto-generated name
                            }
                        } else {
                            //TODO format cell value
                        }

                        if( value != null ) {
                            int columnIndex = SheetDimension.getColumnIndexFromCellRef( cell.getR() );
                            if( columnIndex > 0 ) { // ensures that cell ref is valid
                                columnIndex -= sheetDimension.getFirstColumnIndex();
                                if( columnIndex >= 0 && columnIndex < columnNames.size() ) {
                                    columnNames.set( columnIndex, value );
                                }
                            }
                        }
                    }
                    return columnNames;
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
}
