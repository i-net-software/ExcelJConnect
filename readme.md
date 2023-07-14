[![Build](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml/badge.svg)](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml)

A Excel driver (*.xlsx) written completely in Java (pure Java).

## Usage ##

```java
Connection conn = DriverManager.getConnection( "jdbc:inetexcel:D:\\excel_driver_tests\\column_names.xlsx?hasHeaderRow=false" );
DatabaseMetaData metaData = conn.getMetaData();

// lists all available sheets
try( ResultSet rs = metaData.getProcedures( null, null, null ) ) {
	while( rs.next() ) {
		String sheet = rs.getString( "PROCEDURE_NAME" );
		System.out.println( sheet );
	}
}

// lists all available sheets with their columns 
try( ResultSet rs = metaData.getProcedureColumns( null, null, null, null ) ) {
	while( rs.next() ) {
		String sheet = rs.getString( "PROCEDURE_NAME" );
		String column = rs.getString( "COLUMN_NAME" );
		System.out.println( sheet + " / " + column );
	}
}

// lists all available columns from specified sheet
try( ResultSet rs = metaData.getProcedureColumns( null, null, "SheetName", null ) ) {
	while( rs.next() ) {
		String column = rs.getString( "COLUMN_NAME" );
		System.out.println( column );
	}
}

// reads data from specified sheet
try( CallableStatement stm = conn.prepareCall( "{call SheetName}" ); ResultSet rs = stm.executeQuery() ) {
	while( rs.next() ) {
		String columnValue = rs.getString( "ColumnName" );
		String otherColumnValue = rs.getString( "OtherColumnName" );
		System.out.println( columnValue + ", " + otherColumnValue );
	}
}
```
