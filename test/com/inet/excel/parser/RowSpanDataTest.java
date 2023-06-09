package com.inet.excel.parser;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;

public class RowSpanDataTest {

    @Test
    public void fresh_instance_is_empty() {
        RowSpanData spanData = new RowSpanData();

        assertTrue( spanData.isEmpty() );
        assertEquals( 0, spanData.getFirstColumnIndex() );
        assertEquals( 0, spanData.getLastColumnIndex() );
    }

    @Test
    public void information_about_single_cell_constitutes_valid_row_span() {
        RowSpanData spanData = new RowSpanData();
        spanData.addCellRef( "D2" );
        assertNotEmpty( spanData, 4, 4 );

        spanData = new RowSpanData();
        spanData.addCellRef( "AB1" );
        assertNotEmpty( spanData, 28, 28 );
    }

    @Test
    public void instance_updates_indexes_based_on_collected_span_and_cell_information() {
        RowSpanData spanData = new RowSpanData();

        spanData.addSpanRange( "5:7" );
        assertNotEmpty( spanData, 5, 7 );

        spanData.addSpanRange( "6:8" );
        assertNotEmpty( spanData, 5, 8 );

        spanData.addSpanRange( "4:6" );
        assertNotEmpty( spanData, 4, 8 );

        spanData.addSpanRange( "5:6" );
        assertNotEmpty( spanData, 4, 8 );

        spanData.addCellRef( "D3" );
        spanData.addCellRef( "H9" );
        assertNotEmpty( spanData, 4, 8 );

        spanData.addCellRef( "C5" );
        assertNotEmpty( spanData, 3, 8 );

        spanData.addCellRef( "K1" );
        assertNotEmpty( spanData, 3, 11 );

        spanData.addSpanRange( "2:12" );
        assertNotEmpty( spanData, 2, 12 );

        spanData.addCellRef( "A1" );
        spanData.addCellRef( "AA1" );
        assertNotEmpty( spanData, 1, 27 );
    }

    @Test
    public void instance_ignores_span_and_cell_information_which_are_invalid() {
        RowSpanData spanData = new RowSpanData();

        spanData.addSpanRange( "3:9" );
        assertNotEmpty( spanData, 3, 9 );

        List<String> invalidSpans = Arrays.asList( null, //
                                                   "", //
                                                   "    ", //
                                                   "1", //
                                                   "@", //
                                                   ":", //
                                                   "A:Z", //
                                                   "1:", //
                                                   ":15", //
                                                   "A:15", //
                                                   "1:Z", //
                                                   "0:1", //
                                                   "1:0", //
                                                   "2:1" );

        for( String span : invalidSpans ) {
            spanData.addSpanRange( span );
            assertNotEmpty( spanData, 3, 9, "span=\"" + span + "\"" );
        }

        List<String> invalidCellRefs = Arrays.asList( null, //
                                                      "", //
                                                      "    ", //
                                                      "1", //
                                                      "@", //
                                                      ":", //
                                                      "5A" );

        for( String cellRef : invalidCellRefs ) {
            spanData.addCellRef( cellRef );
            assertNotEmpty( spanData, 3, 9, "cellRef=\"" + cellRef + "\"" );
        }
    }

    private void assertNotEmpty( RowSpanData spanData, int expectedFirstColumnIndex, int expectedLastColumnIndex ) {
        assertNotEmpty( spanData, expectedFirstColumnIndex, expectedLastColumnIndex, "" );
    }

    private void assertNotEmpty( RowSpanData spanData, int expectedFirstColumnIndex, int expectedLastColumnIndex, String msg ) {
        assertFalse( spanData.isEmpty(), msg );
        assertEquals( expectedFirstColumnIndex, spanData.getFirstColumnIndex(), msg );
        assertEquals( expectedLastColumnIndex, spanData.getLastColumnIndex(), msg );
    }
}
