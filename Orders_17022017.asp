<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->


<%

dim rs1 

Set rs1 = Server.CreateObject("ADODB.Recordset")


Dim SortFields
Dim SortOrder
Dim LastSort

SortFields="DATE_OF_ORDER*CATALOGUE_NUMBER*WT*PC"

Dim SortBy
Dim SortArr

SortArr = split(SortFields,"*")

SortBy = ""

if request.QueryString("SortID") = "" then
	if Session("SortID") <> 0 then 
		Session("SortID") = 0
	end if
else
	Session("SortID") = left(request.QueryString("SortID"),2)
end if
	
Session("SortBy") = SortArr(Session("SortID"))

if Session("SortBy") = Session("LastSort") then
	if Session("SortOrder") = "ASC" or Session("SortOrder") = "" then
		Session("SortOrder") = "DESC"
	else
		Session("SortOrder") = "ASC"
	end if
else
	Session("SortOrder") = "ASC"
end if
Session("LastSort") = Session("SortBy")
%>

<%

if request.Form("T1") <> "" then
	response.Write(request.Form("T1"))
end if

%>
<%
if request.QueryString("Supplier") <> "" then
	Session("Supplier") = request.QueryString("Supplier")
end if
if request.QueryString("TPName") <> "" then
	Session("TPName") = request.QueryString("TPName")
end if
if request.QueryString("Customer") <> "" then
	Session("Customer") = request.QueryString("Customer")
end if
if request.QueryString("FLA") <> "" then
	Session("FLA") = request.QueryString("FLA")
end if
if request.QueryString("SLA") <> "" then
	Session("SLA") = request.QueryString("SLA")
end if

%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "A746"
If (Session("Supplier") <> "") Then 
  Recordset1__MMColParam = Session("Supplier")   
End If
%>
<%
Dim Recordset1__MMColParam2
Recordset1__MMColParam2 = "04N09143"
If (Session("Customer")    <> "") Then 
  Recordset1__MMColParam2 = Session("Customer")   
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Collation Screens</title>


<script type="text/javascript">

<!--

function AddData(bc,wt,pc,mi,ma,trid)
{
	var maxItems = 4;
	var maxWeight = 15;
	var maxValue = 350;

	var curItems = 0;
	var curWeight = 0;
	var curValue = 0;
	
	if ( Number(wt) > Number(maxWeight) ) 
		return;
	if ( Number(pc) > Number(maxValue) ) 
		return;
	
	if ( document.forms[0].totalcount.value != '' )
	{
		if ( (Number(document.forms[0].totalcount.value) + 1) > maxItems )
		{
			alert('Exceeded maximum items');
			return;
		}
	}
	if ( document.forms[0].totalweight.value != '' )
	{
		if ( (Number(document.forms[0].totalweight.value) + Number(wt)) > maxWeight )
		{
			alert('Exceeded maximum weight');
			return;
		}
	}
	if ( document.forms[0].totalvalue.value != '' )
	{
		if ( (Number(document.forms[0].totalvalue.value) + Number(pc)) > Number(maxValue) )
		{
			alert('Exceeded maximum value');
			return;
		}
	}

	document.getElementById(trid).bgColor = "#FFFFFF"
	document.getElementById(trid).bgColor = "#99FF99"

	if ( document.forms[0].T1.value == '' )
	{
		document.forms[0].T1.value = bc
		document.forms[0].trackbc.value = bc
		document.forms[0].mi.value = mi
//		document.forms[0].ma.value = ma
	}
	else
	{
		var Contents = document.forms[0].T1.value;
		if ( Contents.indexOf(bc) == -1 ) 
		{
			document.forms[0].T1.value = document.forms[0].T1.value + "\n" + bc
		}
		else
		{
			alert('Item already set to collate');
			return;
		}
	}
	if ( document.forms[0].trackbc.value > bc )
	{
		document.forms[0].trackbc.value = bc;
	}
	

	if ( document.forms[0].totalweight.value == '' )
	{
		document.forms[0].totalweight.value = wt;
	}
	else
	{
		document.forms[0].totalweight.value = Number(document.forms[0].totalweight.value) + Number(wt);
	}

	if ( document.forms[0].totalvalue.value == '' )
	{
		document.forms[0].totalvalue.value = pc;
	}
	else
		document.forms[0].totalvalue.value = Number(document.forms[0].totalvalue.value) + Number(pc);

	if ( document.forms[0].totalcount.value == '' )
	{
		document.forms[0].totalcount.value = "1";
	}
	else
		document.forms[0].totalcount.value = Number(document.forms[0].totalcount.value) + Number(1);
	
	{
		var Items;
		Items = document.forms[0].T1.value.split('\n');
		Items.sort();
		
		var i;
		i = 0;
		while ( i < Items.length )
		{
			var ni;
			ni = Items[i].replace("\n","");
			ni = ni.replace('\n','');
			ni = ni.replace('\r','');
			ni = ni.replace("\r","");
			Items[i] = ni;			
//			alert("[" + ni + "]");
			i = i + 1;
		}		
		
		
		i = 0;
		document.forms[0].T1.value = "";
		while ( i < Items.length )
		{
			if ( Items[i].length > 10 )
			{
//			document.forms[0].T1.value = document.forms[0].T1.value +  Items[i] + "\n";
			if ( document.forms[0].T1.value == '')
			{
				document.forms[0].T1.value = Items[i];
			}
			else
			{
				document.forms[0].T1.value = document.forms[0].T1.value + "\n" + Items[i];
			}
			}
			i = i + 1
		}
//		document.forms[0].T1.value.join('\n');		
	}
	
	
}

function CheckRestrictions()
{
	var maxItems = 4;
	var maxWeight = 6;
	var maxValue = 200;
	if ( document.forms[0].totalcount.value > maxItems )
		document.forms[0].totalcount.style.color = "red";
	else
		document.forms[0].totalcount.style.color = "black";
	if ( document.forms[0].totalvalue.value > maxValue )
		document.forms[0].totalvalue.style.color = "red";
	else
		document.forms[0].totalvalue.style.color = "black";
	if ( document.forms[0].totalweight.value > maxWeight )
		document.forms[0].totalweight.style.color = "red";
	else
		document.forms[0].totalweight.style.color = "black";
}

function ResetData()
{
	document.forms[0].T1.value = "";
	document.forms[0].trackbc.value = "";
	document.forms[0].totalvalue.value = "0";
	document.forms[0].totalweight.value = "0";
	document.forms[0].totalcount.value = "0";
	
	var did = 1;
	
	while ( 1 )
	{
		var el = document.getElementById("tr_id_" + did);
		if (el==null)
		{	
			break;
		}
		el.bgColor = "#E0E0FF";
		did = did + 1;
	}	
}

//-->

</script> 

<link href="css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
	color: #FFFFFF;
	font-weight: bold;
}
.style2 {
	color: #000000;
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
}
.style3 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: small;
}
-->
</style>
</head>

<body bgcolor="E0E0FF" leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
Recordset1.Source = "SELECT *  FROM OrderVw2  WHERE DD_SUPPLIER_CODE = '" + Replace(Recordset1__MMColParam, "'", "''") + "' AND ACCOUNT_REF_NO = '" + Replace(Recordset1__MMColParam2, "'", "''") + "' and DELIVERY_ADDR1='" + replace(Session("FLA"),"'","''") + "' and DELIVERY_ADDR2 ='" + replace(Session("SLA"),"'","''") + "' ORDER BY " + Session("SortBy") + " " + Session("SortOrder")

Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()
if Recordset1.Eof then
  response.Redirect("https://www.ddtrack.co.uk/collation/SuppCust.asp?Supplier=" + Session("Supplier") + "&TPName=" + Session("TPName"))
end if

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>


<table width="100%" border="0" bgcolor="9999FF">
  <tr>
    <td><a href="SuppCust.asp?Supplier=<%=Session("Supplier")%>&TPName=<%=Session("TPName")%>"><img src="Images/ddtrack.gif" width="224" height="42" border="0" /></a></td>
    <td>
      <span class="style1">Trading Partner :</span> <span class="style2"><strong><%=Session("TPName")%>&nbsp;(&nbsp;<%=Session("SUPPLIER")%>&nbsp;)</strong></span><br /><a href="SuppCust.asp?Supplier=<%=Session("Supplier")%>&TPName=<%=Session("TPName")%>" class="style3"></a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<br/>
<table width="100%" border="0" class="MyTable">
  <tr>
    <td width="9%">Customer</td>
    <td width="91%"><%=(Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)%></td>
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
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><img src="Images/green.png" width="16" height="16" />&nbsp;Collated&nbsp;<img src="Images/red.png" width="16" height="16" />&nbsp;Not Collated&nbsp;<img src="Images/black.png" width="16" height="16" />&nbsp;UnCollatable</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>Columns can be sorted by Item, Weight or Value by clicking on the appropriate heading. Note re-sorting will cancel any collation in progress on this screen. </td>
  </tr>
</table>
<br/>
<table border="1" cellpadding="2" cellspacing="0" class="MyTable">
  <tr bgcolor="#9999FF">
    <td>&nbsp;</td>
    <td>Barcode</td>
    <td>Invoice</td>
    <td><a href="Orders.asp?SortID=1">Item</a></td>
    <td>Opt</td>
    <td>Description</td>
    <td><a href="Orders.asp?SortID=2">Weight</a></td>
    <td><a href="Orders.asp?SortID=3">Value</a></td>
    <td><a href="Orders.asp?SortID=0">Date Of Order</a></td>
    <td>Method</td>
    <td>Main Barcode </td>
    <td>Parcel<br />
    Value</td>
    <td>Parcel<br />
      Weight</td>
    <td>Action</td>
  </tr>
  <% 
  dim dynaID
  dynaID = 1
  %>
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
    <tr id="tr_id_<%=dynaID%>"> 
      <td>
		<% if len(trim(Recordset1.Fields.Item("DD_BARCODE_NO").Value)&"") = 16  then %>
			<% if isnull((Recordset1.Fields.Item("DD_PARENT_NO").Value)) then %>
				<img src="Images/red.png" width="16" height="16" />
			<% else %>
				<img src="Images/green.png" width="16" height="16" />
			<% end if 
		else %>
					<img src="Images/black.png" width="16" height="16" />
		<% end if %>	  </td>
      <td><%=(Recordset1.Fields.Item("DD_BARCODE_NO").Value)%></td>
      <td>
		  <% 	
			Dim iv
			iv = Recordset1.Fields.Item("INVOICE_NO").Value
			while ( len(iv) < 5 )
				iv = "0" & iv
			wend 
		  %>	
	  	  <%=iv%>	  </td>
      <td><%=(Recordset1.Fields.Item("CATALOGUE_NUMBER").Value)%></td>
      <td><%=(Recordset1.Fields.Item("OPTION_NUMBER").Value)%></td>
      <td><%=(Recordset1.Fields.Item("ITEM_DESCRIPTION").Value)%></td>
      <td><%=(Recordset1.Fields.Item("WT").Value)%></td>
      <td><%= FormatCurrency((Recordset1.Fields.Item("PC").Value), 2, -2, -2, -2) %></td>
      <td><%
	  	Dim orderDate,orderDate2
		orderDate = Recordset1.Fields.Item("DATE_OF_ORDER").Value
		orderDate2 = mid(orderDate,7,2) & "/" & mid(orderDate,5,2) & "/" & mid(orderDate,1,4)
		%>
      <%=(Recordset1.Fields.Item("DATE_OF_ORDER").Value)%> </td>
      <td>
	  <%=(Recordset1.Fields.Item("LABEL_LEGEND").Value)%>	  </td>
      <td>
	  <% if isnull(Recordset1.Fields.Item("DD_PARENT_NO").Value) then %>
	  	&nbsp;
	  <% else %>
  	    <a href="pdetail.asp?Master=<%=(Recordset1.Fields.Item("DD_PARENT_NO").Value)%>&TPName=<%=Session("TPName")%>"><%=(Recordset1.Fields.Item("DD_PARENT_NO").Value)%></a>
	  <% end if %>
	  	<% 
		dim t1
		dim t2
		dim t3
		t1 =  trim(Recordset1.Fields.Item("DD_BARCODE_NO").Value)
		t2 =  trim(Recordset1.Fields.Item("DD_PARENT_NO").Value) 
		t3 =  trim(Recordset1.Fields.Item("DD_BARCODE_ORIG").Value) 
		
		if t1 = t2  then
		  if t2 = t3 then
		%>
			(M)
		<% end if 
		end if
		%>	  </td>
      <td>
	  	<% if isnull(Recordset1.Fields.Item("C_PC").Value) then %>
			&nbsp;
		<% else %>
	  		<%= FormatCurrency((Recordset1.Fields.Item("C_PC").Value), 2, -2, -2, -2) %>
		<% end if %>	  </td>
      <td>
	  	<% if isnull(Recordset1.Fields.Item("C_WT").Value) then %>
			&nbsp;
		<% else %>
	  	  	<%=(Recordset1.Fields.Item("C_WT").Value)%>
		<% end if %>	  </td>
      <td>
	  	<% 
		Session("ma") = (Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)
		if len(trim(Recordset1.Fields.Item("DD_BARCODE_NO").Value)) = 16  and Recordset1.Fields.Item("WT").Value < 15 and Recordset1.Fields.Item("PC").Value < 350 then
					if len(trim(Recordset1.Fields.Item("DD_PARENT_NO").Value)&"") = 0 then
				%>
<% 
dim MyPrice
MyPrice = 0
on error resume next
if Session("Supplier") = "Y162" then
' force Joe Brown to not take price into account GL Oct 2016 Sponsor : Chris Hammond
	MyPrice = 0
else
	if isnull(Recordset1.Fields.Item("PC").Value) then

 		MyPrice = 0
	else
		MyPrice = Recordset1.Fields.Item("PC").Value
                if err.number <> 0 then
                  MyPrice = 0
                end if
	end if
end if


%>
				
			  			<a href="#" onclick="AddData('<%=trim((Recordset1.Fields.Item("DD_Barcode_NO").Value))%>','<%=(Recordset1.Fields.Item("WT").Value)%>','<%=MyPrice%>','<%=(Recordset1.Fields.Item("INVOICE_NO").Value)%>','<%=(Recordset1.Fields.Item("ACCOUNT_REF_NO").Value)%>','tr_id_<%=dynaID%>')">Collate</a>
			<% else 
			
rs1.ActiveConnection = MM_DD_DB_STRING

rs1.Source = "SELECT IS_COLLATED FROM TP94 where DD_BARCODE_NO='" & Recordset1.Fields.Item("DD_BARCODE_ORIG").Value & "'"

rs1.CursorType = 0
rs1.CursorLocation = 2
rs1.LockType = 1
dim DCAble

DCAble = 1
rs1.open
if not rs1.EOF then
  if rs1.Fields.Item("IS_COLLATED").Value = 1 then
	DCAble = 0
  end if
end if
rs1.Close

			%>
				<% if DCAble then %>
				<a href="dcollate.asp?Barcode=<%=(trim(Recordset1.Fields.Item("DD_BARCODE_NO").Value))%>&Parent=<%=(trim(Recordset1.Fields.Item("DD_PARENT_NO").Value))%>&Supplier=<%=(Recordset1.Fields.Item("DD_SUPPLIER_CODE").Value)%>&TPName=<%=Session("TPName")%>&Orig=<%=(trim(Recordset1.Fields.Item("DD_BARCODE_ORIG").Value))%>">DeCollate</a>
				<% else %>
					Cannot Decollate
				<% end if %>
			<% end if %>
		<% else %>
			UnCollatable
		<% end if %>	  </td>
    </tr>
    <%  
  dynaID = dynaID + 1
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<form id="fm1" name="fm1" method="post" action="collate.asp">

  
 
 

<table width="100%" border="0" class="MyTable">
  <tr>
    <td colspan="2">Collate Items Into A Single Parcel By Clicking on Collate for each line , then check the values below are not exceeding the maximum recommended. Then Click on Commit to store these items . Reset will remove all items so you can start again. <br /></td>
    <td width="3%">&nbsp;</td>
    <td width="3%">&nbsp;</td>
  </tr>
  <tr>
    <td width="27%"><textarea name="T1" rows="10" id="T1" ></textarea></td>
    <td width="67%"><table width="100%" border="0">
      <tr>
        <td width="30%">Main Tracking Barcode </td>
        <td width="27%"><input name="trackbc" type="text" id="trackbc" size="16" maxlength="16" readonly="true" /></td>
        <td width="19%">&nbsp;</td>
        <td width="24%">&nbsp;</td>
      </tr>
      <tr>
        <td>Combined Weight</td>
        <td><input name="totalweight" type="text" id="totalweight" readonly="true"  /></td>
        <td>Max Weight is </td>
        <td> 15 kg </td>
      </tr>
      <tr>
        <td>Combined Value</td>
        <td><input name="totalvalue" type="text" id="totalvalue" disabled="disabled" /></td>
        <td>Max Value is  </td>
        <td>&pound; <% if Session("Supplier") = "Y162" then %> 0 <% else %> 350 <% end if %></td>
      </tr>
      <tr>
        <td>Number of Items </td>
        <td><input name="totalcount" type="text" id="totalcount" disabled="disabled" /></td>
        <td>Max Items </td>
        <td>4</td>
      </tr>
	  <% Session("mi") = "hh" %>
      <tr>
        <td>Master Invoice </td>
        <td><label>
          <input name="mi" type="text" id="mi" readonly="true" />
        </label></td>
        <td><input type="hidden" name="ma" value="<%=Session("ma")%>" />
          <input name="TPName" type="hidden" id="TPName" value="<%=Session("TPName")%>" />
          <input name="SLA" type="hidden" id="SLA" value="<%=Session("SLA")%>" />
          <input name="FLA" type="hidden" id="FLA" value="<%=Session("FLA")%>" /></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><input name="Commit" type="submit" id="Commit" value="Commit" /></td>
        <td><input name="Reset" type="reset" id="Reset" value="Reset" onclick="ResetData()" /></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><input name="Supplier" type="hidden" id="Supplier" value="<%=Session("Supplier")%>" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

</form>
<hr  />

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Set rs1 = Nothing
%>
