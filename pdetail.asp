<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->
<%
Dim Recordset1__Master
Recordset1__Master = "8475844"
If (request.querystring("Master") <> "") Then 
  Recordset1__Master = request.querystring("Master")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
Recordset1.Source = "SELECT *  FROM OrderVw2  WHERE DD_PARENT_NO = '" + Replace(Recordset1__Master, "'", "''") + "'  and DD_SUPPLIER_CODE='" + Session("Supplier") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style2 {color: #0000FF}
.style3 {font-family: Arial, Helvetica, sans-serif}
-->
</style>
</head>

<body>
<% dim AccountRefNo

	AccountRefNo = (Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)
%>
<table width="100%" border="0">
  <tr>
    <td><h4 class="style3"><%=(Recordset1.Fields.Item("DD_SUPPLIER_CODE").Value)%>&nbsp;<%=request.QueryString("TPName")%></h4></td>
    <td>&nbsp;</td>
  </tr>
</table>
<p>
<table width="100%" border="0" class="MyTable">
  <tr>
    <td width="7%">Customer</td>
    <td width="30%"><%=(Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)%></td>
    <td width="63%">Date : <%=Date()%></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><%=(Recordset1.Fields.Item("DELIVERY_ADDR1").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_ADDR2").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_ADDR3").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_ADDR4").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_ADDR5").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_ADDR6").Value)%><br />
        <%=(Recordset1.Fields.Item("DELIVERY_PCODE").Value)%></td>
    <td>&nbsp;</td>
  </tr>
</table>
<h3 class="style1">Parcel Summary</h3>
<p class="style1">Parcel Number : <%=(Recordset1.Fields.Item("DD_PARENT_NO").Value)%></p>
<% 	dim trc 
	trc = 0
%>
<table border="1" cellpadding="3" cellspacing="0" class="MyTable">
  <tr bgcolor="#CCCCCC" >
    <th>Invoice No </th>
    <th>Order Date </th>
    <th>Item</th>
    <th>Amount</th>
    <th>Weight</t>
    <th>Price</th>
    <th>Unique Reference No </th>
  </tr>
  	<% 	Dim tc
  		Dim tw
		Dim tv
		
		tc = 0
		tw = 0
		tv = 0
	%>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
  	<% 	if trc = 0 then 
  			response.write("<TR bgcolor='#FFFFCC'>") 
			trc  = 1
		else 
			response.write("<TR bgcolor='#FFFFFF'>") 
			trc  = 0
		end if 
	%>
      <td><%=(Recordset1.Fields.Item("INVOICE_NO").Value)%></td>
      <td>
      <%=Recordset1.Fields.Item("DATE_OF_ORDER").Value%>	  </td>
      <td><%=(Recordset1.Fields.Item("CATALOGUE_NUMBER").Value)%>&nbsp;/&nbsp;<%=(Recordset1.Fields.Item("OPTION_NUMBER").Value)%></td>
      <td><%=(Recordset1.Fields.Item("QTY").Value)%></td>
      <td><%=(Recordset1.Fields.Item("WT").Value)%></td>
      <td><%= FormatCurrency((Recordset1.Fields.Item("PC").Value), 2, -2, -2, -2) %></td>
      <td><span class="style2"><%=(Recordset1.Fields.Item("DD_BARCODE_NO").Value)%></span></td>
    </tr>
    <tr>
      <td colspan="8"><%=(Recordset1.Fields.Item("ITEM_DESCRIPTION").Value)%>&nbsp;&nbsp;&nbsp;<%=(Recordset1.Fields.Item("OPTION_DESCRIPTION").Value)%></td>
    </tr>
	<%
			tw = tw + (Recordset1.Fields.Item("WT").Value)
		tv = tv + (Recordset1.Fields.Item("PC").Value)
		ti = ti + 1
	%>
    <% 
	
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
    <tr bgcolor="#CCCCCC">
      <th>Total</th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th><%=ti%></th>
      <th><%=tw%></th>
      <th><%= FormatCurrency(tv, 2, -2, -2, -2) %></th>
      <th>&nbsp;</th>
    </tr>
</table>
<h3 class="style1">Returns Information </h3>
<p class="style1">To return any of these items please contact FGH on (0800) 7311731 quoting your account reference no : <%=AccountRefNo%> and your <span class="style2">unique reference number</span> above. </p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
