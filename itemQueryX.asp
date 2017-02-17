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
' # Changes
' # 
' # optimized recordsets as screen took too long to update
' #
' ###############################################################################################
%>


<!--#include file="Connections/DD_DB.asp" -->
<!--#include file="comfunc.asp" -->



<%

Session("txtS") = ""

%>



<%

'response.Write("hello")
'response.Flush()


Dim conn3
Dim rsItemsNotOnStock , rsStockItems
Dim suppCode3, catNum3, optNum3, suppProdCode3, outOfStock3,suppProdCode,itemDesc,optDesc


set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open MM_DD_DB_STRING


Set rsStockItems = Server.CreateObject("ADODB.Recordset")
rsStockItems.ActiveConnection = MM_DD_DB_STRING
rsStockItems.LockType = 3

Set rsItemsNotOnStock = Server.CreateObject("ADODB.Recordset")
rsItemsNotOnStock.ActiveConnection = MM_DD_DB_STRING
rsItemsNotOnStock.LockType = 3

rsItemsNotOnStock.Source = "select db2admin.supplier_items.* from db2admin.supplier_items left outer join stock_items on supplier_items.dd_supplier_code = stock_items.supp_code and supplier_items.catalogue_number = stock_items.cat_no and supplier_items.option_number = stock_items.OPT_NO where supplier_items.dd_supplier_code = '" & Request.QueryString("Supplier") & "' and stock_items.supp_code is null "


rsItemsNotOnStock.Source = "select db2admin.supplier_items.* from db2admin.supplier_items left outer join db2admin.stock_items on db2admin.supplier_items.dd_supplier_code = db2admin.stock_items.supp_code and db2admin.supplier_items.catalogue_number = db2admin.stock_items.cat_no and db2admin.supplier_items.option_number = db2admin.stock_items.OPT_NO where db2admin.supplier_items.dd_supplier_code = '" & Request.QueryString("Supplier") & "' and db2admin.stock_items.supp_code is null "
	
Dim SQLCriteria

rsItemsNotOnStock.Open()
while not rsItemsNotOnStock.EOF 

    suppCode3 = rsItemsNotOnStock.Fields.Item("DD_SUPPLIER_CODE").Value
    catNum3 = rsItemsNotOnStock.Fields.Item("CATALOGUE_NUMBER").Value
    optNum3= rsItemsNotOnStock.Fields.Item("OPTION_NUMBER").Value
	if not isnull(rsItemsNotOnStock.Fields.Item("SUPP_PRODUCT_CODE").Value) then
		suppProdCode = rsItemsNotOnStock.Fields.Item("SUPP_PRODUCT_CODE").Value
		suppProdCode3 = " ='" & replace(rsItemsNotOnStock.Fields.Item("SUPP_PRODUCT_CODE").Value,"'","") & "' "
	else
		suppProdCode3 = " is null "
	end if
		
	
	SQLCriteria = " where supp_code='" & suppCode3  & "' and CAT_NO='" & catNum3 & "' and OPT_NO='" & optNum3 & "' and supp_product_code" & suppProdCode3 
	rsStockItems.Source = "select * from db2admin.stock_items " & SQLCriteria 

	rsStockItems.Open()
	if rsStockItems.EOF then
		' not currently in stock items table so add it . Safe to do a Recordset update here
		rsStockItems.AddNew
		rsStockItems.Fields.Item("CAT_NO").Value     = catNum3
		rsStockItems.Fields.Item("OPT_NO").Value     = optNum3
		rsStockItems.Fields.Item("SUPP_CODE").Value  = suppCode3
		rsStockItems.Fields.Item("ITEM_DESCRIPTION").Value = NZ(rsItemsNotOnStock.Fields.Item("ITEM_DESCRIPTION").Value,null)
		rsStockItems.Fields.Item("OPTION_DESCRIPTION").Value = NZ(rsItemsNotOnStock.Fields.Item("OPTION_DESCRIPTION").Value,null)
		
		rsStockItems.Fields.Item("SUPP_PRODUCT_CODE").Value  = NZ(rsItemsNotOnStock.Fields.Item("SUPP_PRODUCT_CODE").Value,null)
		
		rsStockItems.Update
	end if
	rsStockItems.Close()
	
	' process any due dates for this supplier
	Dim RunSQL
	
	RunSQL = "update stock_items set due_date=null where exists ( select 1 from orders inner join stock_items on stock_items.supp_code = orders.dd_supplier_code and stock_items.cat_no = orders.catalogue_number "
	RunSQL = RunSQL & " and stock_items.opt_no = orders.option_number where stock_items.due_date is not null and stock_items.due_date <= current date  and orders.dd_supplier_code = '" & suppCode3 & "')"
'	response.write(RunSQL)
'	response.Flush()
	conn3.Execute RunSQL
	
	




	
	
	
	
	rsItemsNotOnStock.MoveNext()
wend

rsItemsNotOnStock.Close()

set rsStockItems = Nothing
set rsItemsNotOnStock = Nothing

' now we have done all the work to make sure the orders and stock items are up to date , transfer to the screen specified in
' the URL parameter "Next"


Response.Redirect(Request.QueryString("Next") & ".asp?Supplier=" & Request.QueryString("Supplier") & "&TPName=" & Request.querystring("TPName") & "&RTAG=" & getRTAG())


%>




