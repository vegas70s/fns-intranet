<%
' Script: kvh-ADODB-functions.asp
' Source: KVH Information Systems
' Changes:
'   2014-11-07 jcathcart - created
'
' Procedure:
'
'   GetADODBRecordSet( ConnectionString, SQL )
'
'   @param string ConnectionString - ADODB formatted connection string.
'   @param string SQL - SELECT statement.
'   @return RecordSet 
'
' Description:
'
'   Connects to data source and loads a record set. It can be used to source 
'   data from Excel 2003 compatible spreadsheets using OLEDB.
'   REQUIRES: Microsoft Access Database Engine 2010 Redistributable:
'   http://www.microsoft.com/en-us/download/details.aspx?id=13255'
'
'   Note that the connection must be closed after using the data. See 
'   the function CloseConnection() below.
'
'   Search the Internet for documentation on using VBscript and ADO recordSets.
'
' Example Using Excel Spreadsheet as data source: 
'
' <!--#include file="kvh-ADODB-functions.asp"-->
' < %'
'   ExcelFile = "\\server\share\excelsheet.xls"
'   ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'                      "Data Source=" & ExcelFile & _
'                      ";Extended Properties=""Excel 8.0;HDR=YES;"";" 
'   SQL = "select itemName, itemDescription from [Sheet1$]
'   set RS = GetADODBRecordSet(ConnectionString, SQL)
' % >
'
' ExcelFile = "D:\PortalCOMPLETE\Intranet\admin\fns\fns-cateringform-data.xls"
' SQL = "SELECT [itemCategory] as Category, [itemDesc] as Description, [itemCost] as Cost " & _
'       "FROM [fns_catering$] GROUP BY itemCategory, itemDesc, itemCost"
'
' SQL -- ** DO NOT USE SPACES IN COLUMN NAMES ** 
'     -- Use aliases if necessary, e.g., select [my id] as myid.


function GetADODBRecordSet(ConnectionString, SQL)
' Returns ADODB Record Set.
' Use something like the following in main script: 
' set RS = GetADODBRecordSet(ConnStr,Query)

    set ConnectionObject = Server.CreateObject("ADODB.Connection")
    ConnectionObject.Open ConnectionString

    set RecordSet = Server.CreateObject("ADODB.Recordset")
    RecordSet.Open SQL, ConnectionObject

    set GetADODBRecordSet = RecordSet

end function


function CloseConnection()
'   Call this function after retrieving the data, like from the end of the script.'
    RecordSet.Close
    ConnectionObject.Close
end function

%>
