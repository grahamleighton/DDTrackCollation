<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->

<%

dim rtag

randomize()

rtag = CInt(rnd() * 2000)

%>

<%
'####################Edit javascript####################
%>
<script src="jquery-1.9.1.min.js"></script>

 <script type="text/javascript">
$(document).ready( function() {
   $('#clear').click( function () {
         $('#txtSRUpdate').val(""); 
   });
 });
</script>

<script type="text/javascript">
  $(document).ready(function() {
      $("td:empty").html("&nbsp;");
    });
</script>

<%
'################End of Edit javascrip############
%>


<%
' search functtionality
dim SearchString
'dim goback 
SearchString=""
response.write("RF [" + request.Form + "]<BR>" )
response.write("QS [" + request.QueryString + "]<BR>" )

if request.form("srchSubmit")="Search" then
	txtSearch=request.form("txtSearch")
		
	if  txtSearch = "" then
		Response.Redirect("ItemScreenS.asp?next=ItemScreenS&Supplier=" & request.QueryString("Supplier") & "&TPName=" & request.QueryString("TPName"))
	end if
	
else 
'	txtSearch=request.form("txtSearch")
	txtSearch=request.QueryString("txtSearch")
end if 



if txtSearch <> "" then 
	if len(trim(txtSearch))<> 0 then
			SearchString=" AND (SUPP_PRODUCT_CODE LIKE '%" & txtSearch & "%' OR ITEM_DESCRIPTION LIKE '%" & txtSearch & "%' OR OPTION_DESCRIPTION LIKE '%" & txtSearch & "%')" 
		
	end if		
end if
%>
<%
response.write("TS [" + txtSearch + "]<BR>" )
%>


<%
'sorting function 
	Dim SortFields
	SortFields = "DUE_DATE,SUPP_PRODUCT_CODE,CAT_NO,ITEM_DESCRIPTION,OPTION_DESCRIPTION,DUE_DATE DESC,CAT_NO DESC,SUPP_PRODUCT_CODE DESC,ITEM_DESCRIPTION DESC,OPTION_DESCRIPTION DESC,NEVER_COLLATE,NEVER_COLLATE DESC"
	
	
	Dim SortField
	SortField = split (SortFields,",")
		
'	SortField (0) = "DUE_DATE"
'	SortField(1) = "SUPP_PRODUCT_CODE"
'	SortField (2) = "CAT_NO" 
	
%>


<%
Dim SortFieldToUse 
SortFieldToUse = "CAT_NO" 
If request.querystring("OF") <> "" then
      Dim of_i
	  
    On error resume next
      of_i = CInt(request.querystring("OF"))
    if err.Number <> 0 then 
      of_i = 0
    end if
     if of_i <= UBound(SortField)  then
       SortFieldToUse = SortField(of_i)
	   
    End if
End if
%>


<%
'edit function 
If Request("Save") = "Save" Then

	Dim txtEdit
	txtEdit = Request("txtSRUpdate")
	Dim con
	Dim strSQLCmd
	Set con = Server.CreateObject("ADODB.Connection")
	con.Open MM_DD_DB_STRING
	For Each variableNames in Request.Form
		If Instr(variableNames, "selected") = "1" Then
			Dim prodCode4
			prodCode4 = Split(variableNames, "selected_")(1)		
				
			
			Set rsProd2 = Server.CreateObject("ADODB.Recordset")
			rsProd2.ActiveConnection = MM_DD_DB_STRING
			rsProd2.Source = "SELECT CAT_NO, OPT_NO, SUPP_PRODUCT_CODE FROM DB2ADMIN.STOCK_ITEMS WHERE ID=x'" & prodCode4 & "'"
			rsProd2.Open()

			If Not rsProd2.EOF Then


				Dim suppCode2, catNum2, optNum2, suppProdCode2
				catNum2 = rsProd2.Fields.Item("CAT_NO").Value
				optNum2 = rsProd2.Fields.Item("OPT_NO").Value
				
				
				suppProdCode2 = rsProd2.Fields.Item("SUPP_PRODUCT_CODE").Value
				
				
		if not isNull (rsProd2.Fields.Item("SUPP_PRODUCT_CODE").Value) Then 
				strSQLCmd = "UPDATE DB2ADMIN.ORDERS SET SUPP_PRODUCT_CODE='" & LTRIM(RTRIM(txtEdit)) & "' WHERE CATALOGUE_NUMBER='" & catNum2 & "' AND OPTION_NUMBER='" & optNum2 & "' AND SUPP_PRODUCT_CODE ='" & LTRIM(RTRIM(suppProdCode2)) & "'"			
		
			con.Execute strSQLCmd
		End if 
			
				if isNull (rsProd2.Fields.Item ("SUPP_PRODUCT_CODE").Value) Then
				
					strSQLCmd = "UPDATE DB2ADMIN.ORDERS SET SUPP_PRODUCT_CODE='" & LTRIM(RTRIM(txtEdit)) & "' WHERE CATALOGUE_NUMBER='" & catNum2 & "' AND OPTION_NUMBER='" & optNum2 & "'"
						
				con.Execute strSQLCmd
					end if 
					
			
			if isNull (Request("txtSRUpdate")) Then 
						Dim snull 
						snull = ""
				strSQLCmd = "Update DB2ADMIN.ORDERS SET SUPP_PRODUCT_CODE='" & txtEdit & "' WHERE CATALOGUE_NUMBER='" & catNum2 & "' AND OPTION_NUMBER='" & optNum2 & "'"
				
				con.Execute StrSQLCmd
				
				strSQLCmd = "UPDATE DB2ADMIN.STOCK_ITEMS SET SUPP_PRODUCT_CODE='" & snull & "' WHERE CATALOGUE_NUMBER='"& catNum2 & "' AND OPTION_NUMBER='" & optNum2 & "'"
				end if 
			
			
				response.write("[" & strSQLCmd & "]<BR>")
				response.Flush()
			 			rsProd2.Close()
										
				End If
			
				strSQLCmd = "UPDATE DB2ADMIN.STOCK_ITEMS SET SUPP_PRODUCT_CODE='" & txtEdit & "' WHERE ID=x'" & prodCode4 & "'"
				
				
				response.write("[" & strSQLCmd & "]<BR>")
				response.Flush()
				con.Execute strSQLCmd	
			
					
		End If

	Next
	con.Close
	Set con = Nothing
	
End If

' end of edit
%>
			

<%
'update function
If Request("Update") = "Update" Then
		
	Dim dueDate
	dueDate = Request("txtDate")
	Dim cn
	Dim strSQLCommand
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open MM_DD_DB_STRING
	For Each variableName in Request.Form
		If Instr(variableName, "selected") = "1" Then
			Dim prodCode
			prodCode = Split(variableName, "selected_")(1)

			Dim collateValue
			collateValue = "0"

			
			'For never collate , update orders out_of_stock = 1
			If Request("nevercollate_" & prodCode) = "1" Then
				collateValue = "1"
			End If

			strSQLCommand = "UPDATE DB2ADMIN.STOCK_ITEMS SET "

			'When mark as due date with a date , update ORDERS OUT_OF_STOCK = 1
			If dueDate <> "" Then
				strSQLCommand = strSQLCommand & "DUE_DATE='" & dueDate & "',"
			End If

			strSQLCommand = strSQLCommand & "NEVER_COLLATE=" & collateValue & " WHERE ID=x'" & prodCode & "'"
			
			response.write("[" & strSQLCommand & "]<BR>")
			response.Flush()
			cn.Execute strSQLCommand

            
			Set rsProd = Server.CreateObject("ADODB.Recordset")
			rsProd.ActiveConnection = MM_DD_DB_STRING
			rsProd.Source = "SELECT CAT_NO, OPT_NO, SUPP_PRODUCT_CODE FROM DB2ADMIN.STOCK_ITEMS WHERE ID=x'" & prodCode & "'"
			rsProd.Open()

			If Not rsProd.EOF Then


				Dim suppCode, catNum, optNum, suppProdCode


				catNum = rsProd.Fields.Item("CAT_NO").Value
				optNum = rsProd.Fields.Item("OPT_NO").Value
				suppProdCode = rsProd.Fields.Item("SUPP_PRODUCT_CODE").Value

				strSQLCommand = "UPDATE DB2ADMIN.ORDERS SET OUT_OF_STOCK=" & collateValue & " WHERE CATALOGUE_NUMBER='" & catNum & "' AND OPTION_NUMBER='" & optNum & "' AND SUPP_PRODUCT_CODE ='" & suppProdCode & "'"

				cn.Execute strSQLCommand
				
				
			response.write("[" & strSQLCommand & "]<BR>")
			response.Flush()
					rsProd.MoveNext
			End If

			rsProd.Close()

		End If

	Next
	cn.Close
	Set cn = Nothing
	
	'if txtSearch = txtSearch Then
	'txtSearch = request.QueryString("txtSearch")
	'End if 
	
End If


%>



<%
'clear function 
If Request("Clear") = "Clear" Then
	Dim cn1
	Dim strSQLCommand1
	Set cn1 = Server.CreateObject("ADODB.Connection")
	cn1.Open MM_DD_DB_STRING
	For Each variableName in Request.Form
		If Instr(variableName, "selected") = "1" Then
			Dim prodCode1
			prodCode1 = Split(variableName, "selected_")(1)
				
			Dim collateValue1
			collateValue1 = "0"

			If Request("nevercollate_" & prodCode) = "1" Then
				collateValue1 = "1"
			End If

            strSQLCommand1 = "UPDATE DB2ADMIN.STOCK_ITEMS SET DUE_DATE= Null WHERE ID=x'" & prodCode1 & "'"

			cn1.Execute strSQLCommand1

			Set rsProd = Server.CreateObject("ADODB.Recordset")
			rsProd.ActiveConnection = MM_DD_DB_STRING
			rsProd.Source = "SELECT SUPP_CODE, CAT_NO, OPT_NO, SUPP_PRODUCT_CODE FROM DB2ADMIN.STOCK_ITEMS WHERE ID=x'" & prodCode1 & "'"
			rsProd.Open()

			If Not rsProd.EOF Then


				suppCode = rsProd.Fields.Item("SUPP_CODE").Value
				catNum = rsProd.Fields.Item("CAT_NO").Value
				optNum = rsProd.Fields.Item("OPT_NO").Value
				suppProdCode = rsProd.Fields.Item("SUPP_PRODUCT_CODE").Value
				'When reset update ORDERS OUT_OF_STOCK = 0 
				strSQLCommand1 = "UPDATE DB2ADMIN.ORDERS SET OUT_OF_STOCK=" & collateValue1 & " WHERE DD_SUPPLIER_CODE='" & suppCode & "' AND CATALOGUE_NUMBER='" & catNum & "' AND OPTION_NUMBER='" & optNum & "' AND SUPP_PRODUCT_CODE ='" & suppProdCode & "'"

				cn1.Execute strSQLCommand1
					rsProd.MoveNext
			End If

			rsProd.Close()


		End If
	Next
	cn1.Close
	Set cn1 = Nothing
End If
%>

<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
Recordset1.Source = "SELECT DD_SUPPLIER_CODE, OPTION_DESCRIPTION FROM DB2ADMIN.SUPPLIER_ITEMS"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0

%>

<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.QueryString("Supplier") <> "") Then
  Recordset2__MMColParam = Request.QueryString("Supplier")
End If
%>

<%

'Dim search_text
'Dim search_url
'search_text = request.form("txtSearch")
'if Session("searchparam") = ""  then
'	Session("searchparam") = search_text
'end if

'response.Write("[" + search_text + "]")
'response.Flush()
'response.End()

%>

<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_DD_DB_STRING
Recordset2.Source = "SELECT HEX(ID) AS ID,CAT_NO, DUE_DATE, NEVER_COLLATE, OPT_NO, SELECTED, SUPP_CODE, SUPP_PRODUCT_CODE, OPTION_DESCRIPTION , ITEM_DESCRIPTION FROM DB2ADMIN.STOCK_ITEMS WHERE SUPP_CODE = '" + Replace(Recordset2__MMColParam, "'", "''") +  "'" & SearchString & " Order By " + SortFieldToUse 

Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
'response.Write(Recordset2.Source)
'response.Flush()
'response.End()
Recordset2.Open()


Recordset2_numRows = 0
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_DD_DB_STRING
Recordset3.Source = "SELECT * FROM DB2ADMIN.ORDERS"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
'Recordset3.Open()


'if  Not Recordset3.EOF Then
'			if Not isNull (Recordset2.Fields.Item("DUE_DATE").Value) then
'			 Recordset3.Fields.Item("OUT_OF_STOCK").Value  = 1
'			else if (Reordset2.Fields.Item ("DUE_DATE").Value) Then
'			 Recordset3.Fields.Item ("OUT_OF_STOCK").Value = 0
'			 Recordset3.Update()

'				end if
'			End if
'End if



'Recordset3_numRows = 0
%>

<%
Dim search
Dim search_numRows

Set search = Server.CreateObject("ADODB.Recordset")
search.ActiveConnection = MM_DD_DB_STRING
search.Source = "SELECT SUPP_PRODUCT_CODE, OPTION_DESCRIPTION, ITEM_DESCRIPTION  FROM DB2ADMIN.STOCK_ITEMS  WHERE (SUPP_PRODUCT_CODE LIKE '%'" + search_text + "%' OR ITEM_DESCRIPTION LIKE '%" + search_text + "%' OR OPTION_DESCRIPTION LIKE '%" + search_text + "%')" 
search.CursorType = 0
search.CursorLocation = 2
search.LockType = 1
'search.Open()

search_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
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

	'if txtSearch = "" then 	
  		'	MM_param = Request.QueryString("offset") = 0 
				If (MM_param = "") Then
				 MM_param = Request.QueryString("offset") 
			
			
  			'MM_param = request.QueryString("goback")
  			'MM_param = goback
			'MM_param = Request.QueryString("goback ") 
			'else if ( MM_param = Request.querytxtSearch ) then 
			'MM_param = Request.QueryString("offset") = 0 
				'end if 
		'end if
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then ' last page not a full repeat region
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



Dim URLvars
URLVars = ""
Dim URLStr 
Dim CurrentOffset
CurrentOffset = 0
dim URL_i
URLStr = MM_keepMove
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
	'response.write(CStr(URL_i) + "&nbsp;" + URLVars(URL_i) + "<BR>")
	if left(UCase(URLVars(URL_i)),6) <> "OFFSET" and left(URLVars(URL_i),4) <> "RTAG" and left(ucase(URLVars(URL_i)),9) <> "AMP;TXTSE" and left(ucase(URLVars(URL_i)),5) <> "TXTSE" then
    	if URLStr <> "?" then
        	URLStr = URLStr & "&" & URLVars(URL_i)
        else
            URLStr = URLStr & URLVars(URL_i)
        end if
   end if
   URL_i = URL_i + 1
wend

'response.Write(URLStr + "<BR>")
'response.Write(Request.ServerVariables("URL") + "<BR>")

'MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_urlStr = Request.ServerVariables("URL") & URLStr & MM_moveParam & "="



MM_moveFirst = MM_urlStr & "0"
'move next ..... search functionality
if txtSearch <> "" Then
MM_moveFirst = MM_moveFirst & "&txtSearch=" & txtSearch
end if 



MM_moveLast  = MM_urlStr & "-1"
'move next ..... search functionality
if txtSearch <> "" Then
MM_moveLast = MM_moveLast & "&txtSearch=" & txtSearch
end if 



MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
'move next ..... search functionality
if txtSearch <> "" Then
MM_moveNext = MM_moveNext & "&txtSearch=" & txtSearch
end if




If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If

'prev... search functionality...
if txtSearch <> "" Then
	MM_movePrev = MM_movePrev & "&txtSearch=" & txtSearch
end if 



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Item</title>
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
	
  <link rel="stylesheet" href="http://code.jquery.com/ui/1.9.1/themes/base/jquery-ui.css" />
  <script src="http://code.jquery.com/jquery-1.8.2.js"></script>
   <script src="http://code.jquery.com/ui/1.9.1/jquery-ui.js"></script>
    
	 <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" rel="stylesheet" type="text/css"/>

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
.style30 {
	font-family: Georgia, "Times New Roman", Times, serif;
	font-size: x-small;
}
.style32 {font-family: Georgia, "Times New Roman", Times, serif; font-size: small; }
.style33 {font-size: small}
-->
  </style>    
</head>

<body bgcolor="E0E0FF">



<table width="100%" border="0" bgcolor="9999FF">
  <tr>
    <td><a href="https://www.ddtrack.co.uk/directdespatch.nsf/homepage?readform&amp;TPNAME=<%=replace(request.QueryString("TPName")," ","_")%>&amp;DU=NO&quot;"><img src="Images/ddtrack.gif" width="224" height="42" border="0" /></a></td>
    <td><h2 align="center" class="style28"><span class="style27">Trading Partner: &nbsp;</span> <%=request.QueryString("TPName")%>&nbsp;<%if Recordset2.eof=false then%>
	(<%=(Recordset2.Fields.Item("SUPP_CODE").Value)%>)
	<%end if%></h2></td>
	
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
<br>
<%
Dim OFParam
%>

<table width="1006" height="32" border="0" align="center">
  <tr>
  <td width="534"><p align="left"><span class="style32">Enter Keyword:
  </span>
    <input type="text" name="txtSearch" id="txtSearch" value="" onKeyUp="caps(this)"/>
  <label>
  <input name="srchSubmit" type="submit" id="srchSubmit" value="Search"  />
  </label>
</p></td>
    <td width="462"><span class="style30"><span class="style33">&nbsp;&nbsp;Supplier Reference Update:</span><br />
  &nbsp;
  <input name="txtSRUpdate" type="text" id="txtSRUpdate" value="<%=Request.Form("txtSRUpdate")%>"/>
	<input name="Save" type="submit" id="Save" value="Save" /><button type="button" value="clear" id="clear">Clear</button>
    </span></td>
  </tr>
 
</table><table width="1136" border="1" align="center" cellpadding="2" cellspacing="0" id="Supp" class= "rowHover">
<thead>
  <tr bgcolor="9999FF">
     <td width="68"><div align="center"><span class="style26">TICK</span></div></td>
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

				<a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">CAT / OPT NO </a>		
			</span>	
		</div>	
	</td>






		  <td width="210">
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
							<a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">ITEM DESCRIPTION</a>		   		
							</span>			
			</div>	
	  </td>
 

    <td width="210">
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
			
							
							<a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">
	OPTION DESCRIPTION </a>			
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
	 <a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">SUPPLIER REFERENCE </a>	 	
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

								<a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">NEVER COLLATE</a>
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
					
					
	<a href="ItemScreen.asp?RTAG=<%=rtag%>&OF=<%=OFParam%>&TPName=<%=request.QueryString("TPName")%>&Supplier=<%=request.querystring("Supplier")%>">DUE DATE</a>			
			</span>			
		</div>		
	</td>
 
 
 </tr>
  
  
 
  </thead>
 
  <tbody id="fbody">
  <% While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) %>
    <tr bgcolor="#E0E0FF">
	<td><div align="center">
	    <input <%If (CStr((Recordset2.Fields.Item("SELECTED").Value &"")) = CStr("1")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="selected_<%=Recordset2.Fields.Item("ID").Value%>" type="checkbox" value="selected" onClick="" class="rowselect" />
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
	  <div><%=(Recordset2.Fields.Item("SUPP_PRODUCT_CODE").Value)%></div>
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


</tbody>
</table>


<br />

<br />
<table width="397" border="0" align="center">
    <tr>
      <td width="255"><div align="center">
        <label> <span class="style32">Due Date: </span>
          <input type="text" name="txtDate" id="datepicker" />
          </label>
      </div></td>
      <td width="65"><div align="center">
          <input type="submit" name="Update" value="Update" />
      </div></td>
      <td width="63"><div align="center">
          <input name="Clear" type="submit" id="Clear" value="Clear" />
      </div></td>
    </tr>
  </table>
</form>

<script type="text/javascript">
$(document).ready(function() {
  $('input:text:first').focus();
});
</script>

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
<%
'Recordset3.Close()
Set Recordset3 = Nothing
%>
<%
'search.Close()
Set search = Nothing
%>
