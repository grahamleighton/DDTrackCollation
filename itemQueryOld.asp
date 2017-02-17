<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<%

' ###############################################################################################
' #  DDTrack Rewrite Project                                        Copyright OTTO 2012
' #  
' #  This page has been written as part of the new DDTrack project
' #  itemquery.asp
' #  
' #  This is a non display ASP page that performs the following :
' #  1) Obtains a list of unique items from the ORDERS table for this supplier
' #  2) Using this list transfers the data to the STOCK_ITEMS table either adding or updating
' #  3) Update the ORDERS table to specify whether an order is out of stock. Strictly speaking
' #     this field implies "cannot be collated" more than it is actually out of stock. This is due
' #     to the flag being set that states that this item is to be never collated
' #
' ###############################################################################################
%>


<!--#include file="Connections/DD_DB.asp" -->
<!--#include file="comfunc.asp" -->



<%

Session("txtS") = ""

%>


<%
Dim cn3
Dim rsSupplierItems, rsStockItems
Dim suppCode3, catNum3, optNum3, suppProdCode3, outOfStock3,suppProdCode,itemDesc,optDesc
Dim sqlCommand
Set cn3 = Server.CreateObject("ADODB.Connection")
cn3.Open MM_DD_DB_STRING

Set rsStockItems = Server.CreateObject("ADODB.Recordset")
rsStockItems.ActiveConnection = MM_DD_DB_STRING
rsStockItems.LockType = 3

Set rsSupplierItems = Server.CreateObject("ADODB.Recordset")
rsSupplierItems.ActiveConnection = MM_DD_DB_STRING
rsSupplierItems.Source = "SELECT * FROM DB2ADMIN.SUPPLIER_ITEMS WHERE DD_SUPPLIER_CODE='" & Request.QueryString("Supplier") & "'"
rsSupplierItems.Open()

Dim DescField
' set up an array of all the nullable fields
DescField = split("SUPP_PRODUCT_CODE,ITEM_DESCRIPTION,OPTION_DESCRIPTION",",")
Dim DescIdx

While Not rsSupplierItems.EOF

     
    suppCode3 = rsSupplierItems.Fields.Item("DD_SUPPLIER_CODE").Value
    catNum3 = rsSupplierItems.Fields.Item("CATALOGUE_NUMBER").Value
    optNum3= rsSupplierItems.Fields.Item("OPTION_NUMBER").Value
	if not isnull(rsSupplierItems.Fields.Item("SUPP_PRODUCT_CODE").Value) then
		suppProdCode = rsSupplierItems.Fields.Item("SUPP_PRODUCT_CODE").Value
		suppProdCode3 = " ='" & replace(rsSupplierItems.Fields.Item("SUPP_PRODUCT_CODE").Value,"'","") & "' "
	else
		suppProdCode3 = " is null "
	end if
		

	outOfStock3 = 0
	Dim SQLCriteria
	
	SQLCriteria = " where supp_code='" & suppCode3  & "' and CAT_NO='" & catNum3 & "' and OPT_NO='" & optNum3 & "' and supp_product_code" & suppProdCode3 
	rsStockItems.Source = "select * from db2admin.stock_items " & SQLCriteria 
'	response.Write("opening " + rsStockItems.Source + "<BR>")
	rsStockItems.Open()
	if rsStockItems.EOF then
		' not currently in stock items table so add it . Safe to do a Recordset update here
		rsStockItems.AddNew
		rsStockItems.Fields.Item("CAT_NO").Value     = catNum3
		rsStockItems.Fields.Item("OPT_NO").Value     = optNum3
		rsStockItems.Fields.Item("SUPP_CODE").Value  = suppCode3
		rsStockItems.Fields.Item("ITEM_DESCRIPTION").Value = NZ(rsSupplierItems.Fields.Item("ITEM_DESCRIPTION").Value,null)
		rsStockItems.Fields.Item("OPTION_DESCRIPTION").Value = NZ(rsSupplierItems.Fields.Item("OPTION_DESCRIPTION").Value,null)
		
		rsStockItems.Fields.Item("SUPP_PRODUCT_CODE").Value  = NZ(rsSupplierItems.Fields.Item("SUPP_PRODUCT_CODE").Value,null)
		
		rsStockItems.Update

	else
		' we have a matching item, check that the descriptions are up to date. Need to code a SQL UPDATE for this to work
		' first check the descriptions
		
		DescIdx = 0
		Dim strSQL
		strSQL = "update stock_items set "
		while ( DescIdx < 3 ) 
		
			if ( NZ(rsSupplierItems.Fields.Item(DescField(DescIdx)).Value,null) <> NZ(rsStockItems.Fields.Item(DescField(DescIdx)).Value,null) ) then
				if NZ(rsStockItems.Fields.Item(DescField(DescIdx)).Value,null) <> null then
					strSQL = strSQL & DescField(DescIdx) & "='" & replace(rsStockItems.Fields.Item(DescField(DescIdx)).Value,"'","") & "',"
				end if
			end if
		
			DescIdx = DescIdx + 1
		wend
		strSQL = strSQL & " SUPP_CODE='" & suppCode3 & "' "
		
		' now check the due date as any due date of today or in the past should be cleared
	    If Not IsNull(rsStockItems.Fields.Item("DUE_DATE").Value) Then
			if ( rsStockItems.Fields.Item("DUE_DATE").Value <= Date() ) then
				strSQL = strSQL & ",DUE_DATE=null "
			else
				outOfStock3 = 1
			end if
		end if		
		if NZ(rsStockItems.Fields.Item("NEVER_COLLATE").Value,0) = 1 then
			outOfStock3 = 1
		end if
		
		strSQL = strSQL & SQLCriteria
'		response.Write(strSQL + "<BR>")
		
		' now execute out formed update statement
    	cn3.Execute strSQL
					
    End If
	
	' Now we have either updated or added the STOCK_ITEMS table we need ro cross reference to the ORDERS table 
	' and make sure that any orders reflect the new status
	' the field outOfStock3 will tell us whether this item is available to collate or not
	' value 1 means not collatable
	
	
	strSQL = "update DB2ADMIN.ORDERS set OUT_OF_STOCK=" & outOfStock3 & " where DD_SUPPLIER_CODE='" & suppCode3 & "' and CATALOGUE_NUMBER='" & catNum3 & "' and OPTION_NUMBER='" & optNum3 & "' and SUPP_PRODUCT_CODE" & suppProdCode3 & " and OUT_OF_STOCK<>" & outOfStock3 & " and " & getOrdersStandardCriteria()
	
'	response.Write(strSQL + "<BR>")
		
	' now execute out formed update statement to update the related orders
    cn3.Execute strSQL
				
	
	rsStockItems.Close()
	
	rsSupplierItems.MoveNext()
	
wend
rsSupplierItems.Close()
set rsSupplierItems = Nothing
set rsStockItems = Nothing

' now we have done all the work to make sure the orders and stock items are up to date , transfer to the screen specified in
' the URL parameter "Next"


Response.Redirect(Request.QueryString("Next") & ".asp?Supplier=" & Request.QueryString("Supplier") & "&TPName=" & Request.querystring("TPName") & "&RTAG=" & getRTAG())


%>




