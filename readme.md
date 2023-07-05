[![Build](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml/badge.svg)](https://github.com/i-net-software/ExcelJConnect/actions/workflows/build.yml)

A Excel driver (*.xlsx) written completely in Java (pure Java).

** Usage **

```java
Connection conn = DriverManager.getConnection( "jdbc:inetexcel:{xlsx file}?hasHeaderRow=true" );
DatabaseMetaData metaData = conn.getMetaData();
ResultSet sheets = metaData.getProcedures( null, null, null );
while( rs.next() ) {
    System.out.println( "Sheet: " + rs.getString( "PROCEDURE_NAME" ) );
}
```