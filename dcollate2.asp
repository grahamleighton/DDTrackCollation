<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<!--#include file="Connections/DD_DB.asp" -->
<%

if(request.querystring("Barcode") <> "") then Command1__DDBC = request.querystring("Barcode")

%>
<%


set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_DD_DB_STRING

set rs = Server.CreateObject("ADODB.Recordset")

rs.ActiveConnection = MM_DD_DB_STRING
rs.CursorType = 1
rs.CursorLocation = 2
rs.LockType = 4
if request.QueryString("Orig") = request.QueryString("Parent") then
	rs.Source = "select DD_PARENT_NO,PRIMARY_ITEM_IND,NO_OF_LABELS,DD_BARCODE_NO,ORDER_KEY,INVOICE_NO,ACCOUNT_REF_NO,DD_BARCODE_ORIG   from Orders where DD_PARENT_NO='" & request.QueryString("Parent") & "'"
else
	rs.Source = "select DD_PARENT_NO,PRIMARY_ITEM_IND,NO_OF_LABELS,DD_BARCODE_NO,ORDER_KEY,INVOICE_NO,ACCOUNT_REF_NO,DD_BARCODE_ORIG  from Orders where DD_BARCODE_ORIG='" & request.QueryString("Orig") & "'"
end if

	response.Write(rs.Source)
	response.Write("<BR>")
	
	response.Flush()

Dim NewOrderKey
Dim InvoiceStr
rs.Open()
while not rs.Eof 
	rs.Fields.Item("DD_PARENT_NO").Value = Null
	rs.Fields.Item("PRIMARY_ITEM_IND").Value = "P"
	rs.Fields.Item("NO_OF_LABELS").Value = 1
	rs.Fields.Item("DD_BARCODE_NO").Value = rs.Fields.Item("DD_BARCODE_ORIG").Value 
	NewOrderKey = rtrim(rs.Fields.Item("ACCOUNT_REF_NO").Value)
	InvoiceStr = Right(CStr(100000 +  rs.Fields.Item("INVOICE_NO").Value),5)
	NewOrderKey = NewOrderKey & rs.Fields.Item("INVOICE_NO").Value
	rs.Fields.Item("ORDER_KEY").Value = NewOrderKey
	rs.update
	response.Write("neworderkey = " & NewOrderKey)
	response.Write("updated " & CStr(Err.Number) & "<BR>")
	response.Flush()

	rs.MoveNext

wend
rs.Close
set rs = Nothing
response.Write("Completed")
response.Flush()
response.End()

response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.QueryString("Supplier") + "&TPName=" & request.QueryString("TPName") & "&FLA=" & Session("FLA") & "&SLA=" & Session("SLA"))



if trim(request.QueryString("Barcode")) = trim(request.QueryString("Parent")) and trim(request.QueryString("Barcode")) = trim(request.QueryString("Orig")) Then
	Command1.CommandText = "UPDATE Orders  SET DD_PARENT_NO = NULL,CPC=NULL,CWT=NULL,PRIMARY_ITEM_IND='P',NO_OF_LABELS=1,DD_BARCODE_NO=DD_BARCODE_ORIG WHERE DD_PARENT_NO = '" + Replace(Command1__DDBC, "'", "''") + "' and DD_SUPPLIER_CODE='" + request.querystring("Supplier") + "'"
	Command1.CommandType = 1
	Command1.CommandTimeout = 0
	Command1.Prepared = true
	on error resume next
	Command1.Execute()
	if err.number > 0 then
		response.Write("Error in update<BR>")
		response.Write("Command :<P>")
		response.Write(Command1.CommandText)
		response.Write("<P>")
		response.Write(err.description)
		response.Flush()
		response.End()
	end if
else


	Command1.CommandText = "UPDATE Orders  SET DD_PARENT_NO = NULL,CPC=NULL,CWT=NULL,ORDER_KEY=RTRIM(ACCOUNT_REF_NO)||RIGHT(CAST((100000 + INVOICE_NO) AS char(6)),5),DD_BARCODE_NO=DD_BARCODE_ORIG WHERE DD_BARCODE_ORIG = '" + request.querystring("Orig")  + "' and DD_SUPPLIER_CODE='" + request.querystring("Supplier") + "'"
	Command1.CommandType = 1
	Command1.CommandTimeout = 0
	Command1.Prepared = true
'	response.Write(Command1.CommandText)
'	response.Flush()
	Command1.Execute()

end if


response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.QueryString("Supplier") + "&TPName=" & request.QueryString("TPName") & "&FLA=" & Session("FLA") & "&SLA=" & Session("SLA"))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>
