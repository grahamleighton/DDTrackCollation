<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<%

' ###############################################################################################
' #  DDTrack Rewrite Project                                        Copyright OTTO 2012
' #  
' #  This page has been written as part of the new DDTrack project
' #  Itemscreen.asp
' #  
' #  This routine displays a Supplier's unique items that would be available for collation. It shows
' #  all the supplier's items from the STOCK_ITEMS table and allows the user to specify that it can
' #  never collate this item AND/OR specify a DUE DATE for the item 
' #
' #  Change History
' #  ==============
' #
' #  Graham Leighton 20/11/2012
' #  Initial Build
' ###############################################################################################
%>


<!--#include file="Connections/DD_DB.asp" -->

<!--#include file="comfunc.asp" -->

<%
' Set up a variable to hold how many we will display in the table

Dim RepeatCount
RepeatCount = 10

%>
<%

	' set up some fields so that we can specify the sorting when user clicks on column headings, the sort order is passed
	' as the URL parameter OF and is an integer specifying index into a table defined by SortField

	Dim SortFields
	Dim SortField
	Dim SortFieldToUse 

	SortFields = "DUE_DATE,SUPP_PRODUCT_CODE,CAT_NO@OPT_NO,ITEM_DESCRIPTION,OPTION_DESCRIPTION,DUE_DATE DESC,CAT_NO DESC@OPT_NO DESC,SUPP_PRODUCT_CODE DESC,ITEM_DESCRIPTION DESC,OPTION_DESCRIPTION DESC,NEVER_COLLATE,NEVER_COLLATE DESC"
	
	SortField = split (SortFields,",")
	SortFieldToUse = "CAT_NO" 
	If request.querystring("OF") <> "" then
		Dim of_i
		  
		On error resume next
		  of_i = CInt(request.querystring("OF"))
		if err.Number <> 0 then 
		  of_i = 0
		end if
		 if of_i <= UBound(SortField)  then
		   SortFieldToUse = replace(SortField(of_i),"@",",")
		   
		End if
	End if
%>
<%

If Request("Update") = "Update" Then

	' here the user has clicked on Update so we should be updating the STOCK_ITEMS record with either a positive date or the
	' NEVER_COLLATE flag is to be set
	' this combination is also used to clear the NEVER_COLLATE flag if it was ticked but now the user wants to clear it

	Dim dueDate
	dueDate = Request("txtDate")
	Dim cn
	Dim strSQLCommand
	Dim collateValue
	Dim prodCode
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open MM_DD_DB_STRING
	For Each variableName in Request.Form
		If Instr(variableName, "selected") = "1"  Then
			
			prodCode = Split(variableName, "selected_")(1)

			' default the never_collate flag to be null then see if it has been ticked
			collateValue = "null"

			If Request("nevercollate_" & prodCode) = "1" Then
				' never collate flag has been specified
				collateValue = "1"
			End If

			strSQLCommand = "UPDATE DB2ADMIN.STOCK_ITEMS SET "

			If dueDate <> "" Then
				' a due date has been specified
				strSQLCommand = strSQLCommand & "DUE_DATE='" & dueDate & "',"
			End If

			strSQLCommand = strSQLCommand & "NEVER_COLLATE=" & collateValue & " WHERE ID=x'" & prodCode & "'"

			' update the STOCK_ITEMS table			
			
			cn.Execute strSQLCommand

		End If

	Next
	cn.Close
	Set cn = Nothing
End If

%>
<%
If Request("Clear") = "Clear" Then

	' here the user has pressed the Clear button , this can only really be to clear the due date
	
	Dim cn1
	Dim strSQLCommand1
	Set cn1 = Server.CreateObject("ADODB.Connection")
	cn1.Open MM_DD_DB_STRING
	For Each variableName in Request.Form
		If Instr(variableName, "selected") = "1" Then
			Dim prodCode1
			prodCode1 = Split(variableName, "selected_")(1)

            strSQLCommand1 = "UPDATE DB2ADMIN.STOCK_ITEMS SET DUE_DATE= Null WHERE ID=x'" & prodCode1 & "'"

			cn1.Execute strSQLCommand1

		End If
	Next
	cn1.Close
	Set cn1 = Nothing
End If

%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.QueryString("Supplier") <> "") Then
  Recordset2__MMColParam = Request.QueryString("Supplier")
End If
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_DD_DB_STRING
Recordset2.Source = "SELECT HEX(ID) AS ID,CAT_NO, DUE_DATE, NEVER_COLLATE, OPT_NO, SELECTED, SUPP_CODE, SUPP_PRODUCT_CODE, OPTION_DESCRIPTION , ITEM_DESCRIPTION FROM DB2ADMIN.STOCK_ITEMS WHERE SUPP_CODE = '" + Replace(Recordset2__MMColParam, "'", "''") +  "' order by " + SortFieldToUse 
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
%>
<%
' This will auto update should you use the wizard 
RepeatCount = Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset2_total
Dim Recordset2_first
Dim Recordset2_last


' set the record count
Recordset2_total = Recordset2.RecordCount

' set the number of rows displayed on this page
If (Recordset2_numRows < 0) Then
  Recordset2_numRows = Recordset2_total
Elseif (Recordset2_numRows = 0) Then
  Recordset2_numRows = 1
End If

' set the first and last displayed record
Recordset2_first = 1
Recordset2_last  = Recordset2_first + Recordset2_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset2_total <> -1) Then
  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
  End If
  If (Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset2_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset2_total=0
  While (Not Recordset2.EOF)
    Recordset2_total = Recordset2_total + 1
    Recordset2.MoveNext
  Wend
'  response.Write(CStr(Recordset2_total) + "<BR>")

  ' reset the cursor to the beginning
  If (Recordset2.CursorType > 0) Then
    Recordset2.MoveFirst
  Else
    Recordset2.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset2_numRows < 0 Or Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If

  ' set the first and last displayed record
  Recordset2_first = 1
  Recordset2_last = Recordset2_first + Recordset2_numRows - 1

  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
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

Set MM_rs    = Recordset2
MM_rsCount   = Recordset2_total
MM_size      = Recordset2_numRows
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
<title>Item Maintenance For <%=request.querystring("Supplier")%></title>
<style type="text/css">
<!--
.style2 {font-family: Verdana, Arial, Helvetica, sans-serif}
-->


</style>

</style>

<style type="text/css">
table.rowHover tr:hover
	{ background-color: #ccc;
		cursor: pointer;
	 }
table.rowHover tr:first-child:hover {
    background-color:#9999FF;
}





</style>
	
  <link rel="stylesheet" href="jquery-ui.css" />
  <script src="jquery-1.8.2.js"></script>
   <script src="jquery-ui.js"></script>

   <script>
     $(function() {
	    $( "#datepicker" ).datepicker({dateFormat: 'dd/mm/yy'});
   });    </script>

 <script>
 $(document).ready(function() {
        $('#Supp tr').click(function(event) {
          if (event.target.type !== 'checkbox') {
            $('.rowselect', this).trigger('click');
          }
        });
});

 $(document).ready(function() {
        $('.never').click(function(event) {
            $('.rowselect', $(this).parents('tr')).prop('checked', true);

        });
});

</script>
  <style type="text/css">
<!--
.style26 {font-family: "Courier New", Courier, monospace}
.style27 {color: #FFFFFF}
.style28 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
}
-->
  </style>

</head>

<body bgcolor="E0E0FF">

<table width="100%" border="0" bgcolor="9999FF">
  <tr>
    <td><a href="https://www.ddtrack.co.uk/directdespatch.nsf/homepage?readform&amp;TPNAME=<%=replace(request.QueryString("TPName")," ","_")%>&amp;DU=NO"><img src="Images/ddtrack.gif" width="224" height="42" border="0" /></a></td>
    <td><h2 align="center" class="style28"><span class="style27">Trading Partner: &nbsp;</span> <%=request.QueryString("TPName")%>&nbsp;(<%=Request.QueryString("Supplier")%>)</h2></td>


    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<form id="form1" name="form1" method="post" action="">
  <p class="style2">This screen allows you to advise  which items are temporarily unavailable for collation.</p>
  <p class="style2">Please select the relevant  catalogue number(s), enter the date when the product will be available and  click &ldquo;Update&rdquo;.</p>
  <table border="0" width="50%" align="center">
  <tr>
    <td width="23%" align="center"><% If MM_offset <> 0 Then %>
          <a href="<%=MM_moveFirst%>"><img src="First.gif" border=0></a>
          <% End If ' end MM_offset <> 0 %>    </td>
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

<%
Dim OFParam

%>

<table width="1136" border="1" align="center" cellpadding="2" cellspacing="0" id="Supp" class= "rowHover">
  <tr bgcolor="9999FF">
     <td width="68"><div align="center"><span class="style26">TICK</span></div>
       <label>
       <div align="center"></div>
      </label></td>
    <td width="190">
		<div align="center">
			<span class="style26">

<%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 2
	if instr(SortFieldToUse,"CAT_NO" ) > 0 then
		if ( instr(SortFieldToUse,"DESC") ) > 0 then
			OFParam = 2
		else
			OFParam = 6		
		end if
	end if
%>

				<a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">CAT /<BR/>OPT NO </a>		
			</span>	
		</div>	
	</td>






		  <td width="280">
		  			<div align="center">
							<span class="style26">
							<%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 3
	if instr(SortFieldToUse,"ITEM_DESCRIPTION" ) > 0 then
		if ( instr(SortFieldToUse," DESC") ) > 0 then
			OFParam = 3
		else
			OFParam = 8
		end if
	end if
%>
							<a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">ITEM<BR/>DESCRIPTION</a>		   		
							</span>			
			</div>	
	  </td>
 

    <td width="120">
				<div align="center">
							<span class="style26"> 
						
<%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 4
	if instr(SortFieldToUse,"OPTION_DESCRIPTION" ) > 0 then
		if ( instr(SortFieldToUse," DESC") ) > 0 then
			OFParam = 4
		else
			OFParam = 9
		end if
	end if
%>
			
							
							<a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">
	OPTION<BR/>DESCRIPTION </a>			
		</span>		
	</div>
</td>
	
	
	 <td width="300">
	 				<div align="center">
										<span class="style26">
										
										
<%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 1
	if instr(SortFieldToUse,"SUPP_PRODUCT_CODE" ) > 0 then
		if ( instr(SortFieldToUse,"DESC") ) > 0 then
			OFParam = 1
		else
			OFParam = 7		
		end if
	end if
%>
	 <a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">SUPPLIER<BR />REFERENCE </a>	 	
	 	</span>	
	</div>	
</td>
	 
	 
	 
	 
	 
	 
	  <td width="120">
	 			 <div align="center">
	 								 <span class="style26">
									 
									 <%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 10
	if instr(SortFieldToUse,"NEVER_COLLATE" ) > 0 then
		if ( instr(SortFieldToUse,"DESC") ) > 0 then
			OFParam = 10
		else
			OFParam = 11		
		end if
	end if
%>

								<a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">NEVER COLLATE</a>
								</span>
		</div>
	  </td>
  
  
  
   <td width="210">
   			<div align="center">
					<span class="style26">
					
					<%
'	look at sortfieldtouse , decide what OF should be
'	if sortfieldtouse relates to this column , work out is it asc or desc and assign OF accordingly
	
	OFParam = 0
	if instr(SortFieldToUse,"DUE_DATE" ) > 0 then
		if ( instr(SortFieldToUse,"DESC") ) > 0 then
			OFParam = 0
		else
			OFParam = 5		
		end if
	end if
%>
					
					
	<a href="ItemScreen.asp?RTAG=<%=getRTAG()%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">DUE DATE</a>			
			</span>			
		</div>		
	</td>
 </tr>
  
  
  
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) %>
    <tr bgcolor="#E0E0FF">
	<td><div align="center">
	    <input <%If (CStr((Recordset2.Fields.Item("SELECTED").Value &"")) = CStr("1")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="selected_<%=Recordset2.Fields.Item("ID").Value%>" type="checkbox" value="selected" onclick="" class="rowselect" />
	  </div></td>

      <td><div align="center"><span class="style26"><%=(Recordset2.Fields.Item("CAT_NO").Value)%> / <%=(Recordset2.Fields.Item("OPT_NO").Value)%></span></div></td>



	      <td><div align="center"><span class="style26"><% If isNull (Recordset2.Fields.Item("ITEM_DESCRIPTION").Value)then %> &nbsp;
	  <%else%>
	 <%=(Recordset2.Fields.Item("ITEM_DESCRIPTION").Value)%>
	  <%end if%>
	  </span></div></td>





      <td><div align="center"><span class="style26"><% If isNull (Recordset2.Fields.Item("OPTION_DESCRIPTION").Value) then %>&nbsp;
	  <%else%>
	  <%=(Recordset2.Fields.Item("OPTION_DESCRIPTION").Value)%>
	  <%end if%>
	  </span></div></td>


	  <td><div align="center"><span class="style26"><%If isnull (Recordset2.Fields.Item("SUPP_PRODUCT_CODE").Value) then %> &nbsp;
	  <%else%>
	  <%=(Recordset2.Fields.Item("SUPP_PRODUCT_CODE").Value)%>
	  <%end if%>
	  </span></div></td>



	  <td><div align="center">
          <input type="checkbox" name="nevercollate_<%=(Recordset2.Fields.Item("ID").Value)%>" <%If Trim(CStr(Recordset2.Fields.Item("NEVER_COLLATE").Value &"")) = Trim("1") Then Response.Write("checked=""checked""") : Response.Write("")%> value="1" class="never"/>
      </div></td>
      <td><div align="center"><span class="style26"><% If isNull (Recordset2.Fields.Item("DUE_DATE").Value) then %> &nbsp;
	  <%else%>
	  <%=(Recordset2.Fields.Item("DUE_DATE").Value)%>
	  <%end if%>
	  </span></div></td>
    </tr>
    <%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset2.MoveNext()
Wend
%>
</table>


<p align="center">

<%

dim PageNum
dim recCount

PageNum = 1
recCount = 1
Dim URLvars
URLVars = ""
Dim URLStr 
Dim CurrentOffset
CurrentOffset = 0
dim URL_i
URLStr = request.QueryString
on error resume next
CurrentOffset = CInt(request.QueryString("Offset"))
if err.number <> 0 then
	CurrentOffset = 0
end if
on error goto 0



URLVars = split(URLStr , "&")
URL_i = 0
URLStr = "?"
while ( URL_i <= UBound(URLVars))
	if left(UCase(URLVars(URL_i)),6) <> "OFFSET" and left(URLVars(URL_i),4) <> "RTAG" then
		if URLStr <> "?" then
			URLStr = URLStr & "&" & URLVars(URL_i)
		else
			URLStr = URLStr & URLVars(URL_i)
		end if
	end if
	URL_i = URL_i + 1
wend



if ( Recordset2_total > RepeatCount ) then

	while ( recCount < Recordset2_total ) 
	
		if ((PageNum-1) * RepeatCount) = CurrentOffset then
		%>
		
		<%=PageNum%>&nbsp;
	    <% else %>
				<a href="<%=Request.ServerVariables("URL")%><%=URLStr%>&RTAG=<%=getRTAG()%>&Offset=<%=((PageNum-1)*RepeatCount)%>"><%=PageNum%></a>&nbsp;
	    
		<% end if
		if ( PageNum Mod 40 = 0 ) then
			response.Write("<BR>")
		end if
			
		recCount = recCount + RepeatCount
		PageNum = PageNum + 1
	wend

end if

%>
</p>

<table width="793" border="0" align="center">

	<% if Recordset2_total > 0 then %>
    <tr>
      <td width="383"><div align="center">
        <label> <span class="style26">Due Date: </span>
          <input type="text" name="txtDate" id="datepicker" />
          </label>
      </div></td>
      <td width="233"><div align="right">
          <input type="submit" name="Update" value="Update" />
      </div></td>
      <td width="163"><div align="right">
          <input name="Clear" type="submit" id="Clear" value="Clear" />
      </div></td>
    </tr>
	<% else %>
		<tr>
		<td class="style2" colspan="3" align="center">There are no items available</td>
		</tr>
	<% end if %>
  </table>
</form>

</body>
</html>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>