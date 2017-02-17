<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("Supplier") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("Supplier")
End If
%>


<%

Dim CustSearch
CustSearch = ""
if request.form("fromForm") = "Y" then
  CustSearch = ucase(request.form("txtSearch"))
end if 

%>


<%
Dim SuppCustCountsStub
Dim SuppCustCountsSQL 
Dim SuppCustAccountsSQL
Dim SuppGroupBy


SuppCustCountsStub = "SELECT DISTINCT DD_SUPPLIER_CODE, ACCOUNT_REF_NO, DELIVERY_ADDR1, TP94_TSTAMP, COUNT(*) AS CT, DELIVERY_ADDR2 FROM "  
SuppCustCountsStub = SuppCustCountsStub & " ( SELECT orders.* ,CAST(WEIGHT AS FLOAT) AS WT ,CAST(PRICE AS FLOAT) AS PC FROM Orders "
SuppCustCountsStub = SuppCustCountsStub & " inner join stock_items on orders.dd_supplier_code = stock_items.supp_code and orders.catalogue_number = stock_items.cat_no and orders.option_number = stock_items.opt_no "
SuppCustCountsStub = SuppCustCountsStub & " WHERE ( out_of_stock=0 or out_of_stock is null ) and NO_OF_LABELS=1 and   (TP94_TSTAMP IS NULL) and ( TP9A_TSTAMP is null ) "
SuppCustCountsStub = SuppCustCountsStub &  " and tp91_tstamp is not null and tp93_tstamp is null and tp9a_tstamp is null and tp9b_tstamp is null and tp9k_tstamp is null "
SuppCustCountsStub = SuppCustCountsStub & " and tp9d_tstamp is null and CAST(WEIGHT AS FLOAT) < 15 and CAST(PRICE AS FLOAT) < 350 "
SuppCustCountsStub = SuppCustCountsStub & " and  dd_supplier_code='<<SUPPLIERCODE>>' and coalesce(stock_items.never_collate,0) <> 1 "
SuppCustCountsStub = SuppCustCountsStub & ") as aa "

SuppGroupBy = " GROUP BY DD_SUPPLIER_CODE, ACCOUNT_REF_NO, DELIVERY_ADDR1, TP94_TSTAMP, DELIVERY_ADDR2  HAVING (COUNT(*) > 1) "

SuppCustCountsSQL = replace(SuppCustCountsStub, "<<SUPPLIERCODE>>" , Recordset1__MMColParam ) & "  WHERE DD_PARENT_NO IS NULL  " & SuppGroupBY

SuppCustAccountsSQL = replace(SuppCustCountsStub, "<<SUPPLIERCODE>>" , Recordset1__MMColParam ) & " WHERE DD_PARENT_NO IS NULL  and (ACCOUNT_REF_NO like '" + CustSearch + "%' or DELIVERY_ADDR1 like '%" + CustSearch + "%')" & SuppGroupBy


%>


<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
'Recordset1.Source = "SELECT *  FROM SuppCustCounts  WHERE DD_Supplier_CODE = '" + Replace(Recordset1__MMColParam, "'", "''") + "'"
Recordset1.Source = SuppCustCountsSQL
if CustSearch <> "" then
'Recordset1.Source = "SELECT *  FROM SuppCustCounts  WHERE DD_Supplier_CODE = '" + Replace(Recordset1__MMColParam, "'", "''") + "' and (ACCOUNT_REF_NO like '" + CustSearch + "%' or DELIVERY_ADDR1 like '%" + CustSearch + "%')"
Recordset1.Source = SuppCustAccountsSQL
  
end if

Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
' response.Write(Recordset1.Source)
' response.Flush()
'response.End()
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_DD_DB_STRING
Recordset2.Source = "SELECT *  FROM DB2ADMIN.COLLATED WHERE DD_Supplier_CODE = '" + Replace(Recordset1__MMColParam, "'", "''") + "'"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 25
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset1_total
Dim Recordset1_first
Dim Recordset1_last

' set the record count
Recordset1_total = Recordset1.RecordCount

' set the number of rows displayed on this page
If (Recordset1_numRows < 0) Then
  Recordset1_numRows = Recordset1_total
Elseif (Recordset1_numRows = 0) Then
  Recordset1_numRows = 1
End If

' set the first and last displayed record
Recordset1_first = 1
Recordset1_last  = Recordset1_first + Recordset1_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset1_total <> -1) Then
  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If
  If (Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset1
MM_rsCount   = Recordset1_total
MM_size      = Recordset1_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset2_first = MM_offset + 1
Recordset2_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset2_first > MM_rsCount) Then
    Recordset2_first = MM_rsCount
  End If
  If (Recordset2_last > MM_rsCount) Then
    Recordset2_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Collatable Parcels</title>
<link href="css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
	color: #FFFFFF;
	font-weight: bold;
}
.style4 {
	color: #000000;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: small;
}
.style5 {font-family: Arial, Helvetica, sans-serif}
-->
</style>
</head>

<body bgcolor="E0E0FF" leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table width="100%" border="0" bgcolor="9999FF">
  <tr>
    <td><a href="https://www.ddtrack.co.uk/directdespatch.nsf/homepage?readform&amp;TPNAME=<%=replace(request.QueryString("TPName")," ","_")%>&amp;DU=NO"><img src="Images/ddtrack.gif" width="224" height="42" border="0" /></a></td>
    <td><span class="style2">Trading Partner :</span> <span class="style4"><%=request.QueryString("TPName")%>&nbsp;(&nbsp;<%=Request.QueryString("Supplier")%>&nbsp;)<br />
     </span></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<p class="style1">Below is a list of Customers that have more than one parcel, these could be collated where possible </p>
<table border="0" width="50%" align="center">
  <tr>
    <td width="23%" align="center"><% If MM_offset <> 0 Then %>
        <a href="<%=MM_moveFirst%>"><img src="First.gif" border=0></a>
        <% End If ' end MM_offset <> 0 %>
</td>
    <td width="31%" align="center"><% If MM_offset <> 0 Then %>
          <a href="<%=MM_movePrev%>"><img src="Previous.gif" border=0></a>
          <% End If ' end MM_offset <> 0 %>    </td>
    <td width="23%" align="center"><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveNext%>"><img src="Next.gif" border=0></a>
          <% End If ' end Not MM_atTotal %>    </td>
    <td width="23%" align="center"><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveLast%>"><img src="Last.gif" border=0></a>
          <% End If ' end Not MM_atTotal %>    </td>
  </tr>
</table>

<form id="form1" name="form1" method="post" action="SuppCust.asp?<%=request.querystring%>">
  <table width="481" border="0" cellpadding="3" cellspacing="0">
    <tr>
      <th width="181" class="style1" scope="row">Search Customer </th>
      <td width="168"><label>
        <input name="txtSearch" type="text" class="style1" id="txtSearch" />
      </label></td>
      <td width="114"><label>
        <input type="submit" name="Submit" value="Search" />
        <input name="fromForm" type="hidden" id="fromForm" value="Y" />
      </label></td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>
  <%  if Recordset1.Eof then %>
  
  <span class="style5">No Customers have any items that can be collated</span>
  <% else %>
  
</p>
<table cellpadding="3" cellspacing="0" class="MyTable">
  <tr bgcolor="#9999FF">
    <td width=100>Customer No </td>
    <td width=4>&nbsp;</td>
    <td width=200>Name</td>
    <td width=200>First Line Address</td>
    <td width=100>No Of Parcels </td>
  </tr>
	<tr>
	  	<td height="5" colspan="5">&nbsp;</td>
	</tr>

  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr>
      <td  background="/graphics/reliefbg.gif"><div align="right"><a href="Orders.asp?Supplier=<%=(Recordset1.Fields.Item("DD_SUPPLIER_CODE").Value)%>&amp;Customer=<%=(Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)%>&amp;TPName=<%=request.QueryString("TPName")%>&amp;FLA=<%=Server.URLEncode((Recordset1.Fields.Item("DELIVERY_ADDR1").Value))%>&amp;SLA=<%=Server.URLEncode((Recordset1.Fields.Item("DELIVERY_ADDR2").Value))%>" class="style5"><%=(Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)%></a></div></td>
	  <td width=4 background="graphics/reliefright.gif">&nbsp;</td> 
      <td><span class="style5"><%=(Recordset1.Fields.Item("DELIVERY_ADDR1").Value)%></span></td>
      <td><span class="style5">
      <%=(Recordset1.Fields.Item("DELIVERY_ADDR2").Value)%></span></td>
      <td><div align="center" class="style5"><%=(Recordset1.Fields.Item("CT").Value)%></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<p>
  <% end if %>
</p>
<hr />
<p class="style5">Orders that have already been collated together by Customer </p>
<% if Recordset2.Eof then %>
	<span class="style5">No items have been collated yet</span>
<% else %>
<table cellpadding="3" cellspacing="0" class="MyTable">
  <tr bgcolor="#9999FF">
    <td width="100">Customer No </td>
    <td width=4>&nbsp;</td>
    <td width="200">Name </td>
    <td width="200">First Line Address </td>
    <td width="100">No Of Parcels </td>
  </tr>
	<tr>
	  	<td height="5" colspan="5">&nbsp;</td>
	</tr>
  <% While ((Repeat2__numRows <> 0) AND (NOT Recordset2.EOF)) %>
    <tr>
      <td background="graphics/reliefbg.gif"><div align="right"><a href="Orders.asp?Supplier=<%=(Recordset2.Fields.Item("DD_SUPPLIER_CODE").Value)%>&amp;Customer=<%=(Recordset2.Fields.Item("ACCOUNT_REF_NO").Value)%>&amp;TPName=<%=request.QueryString("TPName")%>&amp;FLA=<%=(Recordset2.Fields.Item("DELIVERY_ADDR1").Value)%>&amp;SLA=<%=(Recordset2.Fields.Item("DELIVERY_ADDR2").Value)%>" class="style5"><%=(Recordset2.Fields.Item("ACCOUNT_REF_NO").Value)%></a></div></td>
      <td width=4 background="graphics/reliefright.gif">&nbsp;</td>
      <td><span class="style5"><%=(Recordset2.Fields.Item("DELIVERY_ADDR1").Value)%></span></td>
      <td><span class="style5"><%=(Recordset2.Fields.Item("DELIVERY_ADDR2").Value)%></span></td>
      <td><div align="center" class="style5"><%=(Recordset2.Fields.Item("CT").Value)%></div></td>
    </tr>
	
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  Recordset2.MoveNext()
Wend
%>
</table>
<% end if %>
<hr  />
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
