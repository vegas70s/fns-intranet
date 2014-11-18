<%@LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% 'This #include loads the database function used below. %>
<!--#include file="kvh-ADODB-functions.asp"-->
<%

    dim ConnectionObject
    dim RecordSet

    dim ExcelFile
    dim ConnectionString
    dim SQL
    dim vbCRLFTB
    dim RS
    dim Field
    dim RowID
    dim RowNum
    dim ItemCost
    dim CategoryPrevious
    dim TableData

' ExcelFile is used in part of the ConnectionString.
ExcelFile = "D:\PortalCOMPLETE\Intranet\admin\fns\fns-cateringform-data.xls"

' ConnectionString is passed to the DB function.
ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & ExcelFile & _
                   ";Extended Properties=""Excel 8.0;HDR=YES;"";" 

' SQL is the query passed to the DB function.
SQL = "SELECT [itemCategory] AS Category, " & _
             "[itemDesc] AS Description, " & _
             "[itemCost] AS Cost " & _
       "FROM [fns_catering$] " & _
      "WHERE RTRIM([itemDesc]) NOT LIKE '' " & _
        "AND RTRIM([itemCategory]) " & _
        "AND ISNUMERIC([itemCost])"

'Call the DB function to fetch the Record Set. Function is included above.
set RS = GetADODBRecordSet(ConnectionString,SQL)

'New Line and a four-space "tab" to make the source code pretty.
vbCRLFTB = vbCRLF & Space(4) 

%><!DOCTYPE html>
<html>
<head>
	<title>FNS Catering Request</title>
    <link rel="stylesheet" href="css/meyer-reset.css" />
    <link rel="stylesheet" href="css/kvh-fns-catering.css" />
</head>
<body>

<form 
    name="formcatering" 
    action="kvh-fns-catering-process.asp"
    method="post" >

<div>

<table id="itemtable" >
<thead>

<tr><% 'Table Headers'
for each Field in RS.Fields
	response.write vbCRLFTB & "<th>" & Field.Name & "</th>" 
next
%>
    <th>Item<br/>Qty</th>
    <th>Item<br/>Total</th>
    <th>Add<br/>Remove</th>
</tr>

</thead>
<tbody>

<% 'Table Rows'
while not RS.eof

    RowNum = RowNum + 1
    RowID = cStr(RowNum)
    if RowNum < 10 then RowID = "0" & RowID 

		for each Field in RS.Fields

            if Field.Name = "Category" then
                'If new Category, DisplayCategoryRow'
                if CategoryPrevious <> Field.Value then
                    CategoryPrevious = Field.Value
                    TableData = "<tr class=""itemcategory"">" & vbCRLFTB & _
                        "<td colspan=" & RS.Fields.Count + 3 & ">" & Field.Value  & _
                        "</td>" & vbCRLF & "</tr>" & vbCRLF & _
                        "<tr class=""item"">" & vbCRLFTB & "<td>&nbsp;</td>"
                else
                    TableData = "<tr class=""item"">" & vbCRLFTB & "<td>&nbsp;</td>"
                end if


            elseif Field.Name = "Cost" then
                ItemCost = FormatNumber(Field.Value,2)
                TableData = vbCRLFTB & _
                    "<td><input name=""itemcost" & RowID & """ size=3 value=""" & ItemCost & """ ></td>" & _
                    vbCRLFTB &_
                    "<td><input name=""itemqty" & RowID & """ size=3 value="""" ></td>" & _
                    vbCRLFTB & _
                     "<td><input name=""itemtot" & RowID & """ size=3 value=""0.00""></td>" & _
                    vbCRLFTB & _
                    "<td><input type=""button"" " & _
                        "name=""itemadd" & RowID & """ value=""+""><input type=""button"" " & _
                        "name=""itemdel" & RowID & """ value=""-""></td>"
            else
                TableData = "<td>" & Field.Value & "</td>"
            end if

            response.write TableData

		next

	response.write vbCRLF & "</tr>" & vbCRLF
	RS.movenext

wend

CloseConnection()

'%>
</tbody>
</table>



</div>
</form>
</body>
</html>