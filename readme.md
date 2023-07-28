[![Build](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml/badge.svg)](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml)
[![Maven](https://img.shields.io/maven-central/v/de.inetsoftware/exceljconnect.svg)](https://mvnrepository.com/artifact/de.inetsoftware/exceljconnect)
[![JitPack](https://jitpack.io/v/i-net-software/ExcelJConnect.svg)](https://jitpack.io/#i-net-software/ExcelJConnect/master-SNAPSHOT)

A Excel driver (*.xlsx) written completely in Java (pure Java).

## Dependencies ##
No dependencies to other libraries are needed.

```
repositories {
    mavenCentral()
}

dependencies {
    implementation 'de.inetsoftware:exceljconnect:+'
}
```

## Usage ##

```java
Connection conn = DriverManager.getConnection( "jdbc:inetexcel:{xlsx file}?hasHeaderRow=false" );
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
