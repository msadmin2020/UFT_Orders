'SystemUtil.Run "E:\Orders\MSVC\Release\Orders.exe"

Dim RowCount, i, d_StartDate, d_ExpiryDate, iColCount, iRowNum,sListValue, iRow
RowCount=Datatable.GetRowCount
For i = 1 To RowCount
    Print(i)
	Datatable.SetCurrentRow(i)
	
	'Convert String to date format while accessing from datatable
	d_StartDate=CDate(DataTable("sStartDate", dtGlobalSheet))
    d_ExpiryDate=CDate(DataTable("sExpiryDate", dtGlobalSheet))
  	
	RunAction "OpenNewOrder [SingleOrder]", oneIteration
    
    Window("Orders").Dialog("New order").WinComboBox("Product :").Select DataTable("sProduct", dtGlobalSheet) @@ hightlight id_;_3868960_;_script infofile_;_ZIP::ssf4.xml_;_
    
   'Type keys to set the StartDate @@ hightlight id_;_787740_;_script infofile_;_ZIP::ssf5.xml_;_
    Window("Orders").Dialog("New order").WinCalendar("StartDate").Click
    Window("Orders").Dialog("New order").WinCalendar("StartDate").Type(d_StartDate)
    
    Window("Orders").Dialog("New order").WinEdit("txt_Customer Name:").Set DataTable("sCustomerName", dtGlobalSheet) @@ hightlight id_;_2098116_;_script infofile_;_ZIP::ssf6.xml_;_
    Window("Orders").Dialog("New order").WinEdit("txt_Street:").Set DataTable("sStreet", dtGlobalSheet) @@ hightlight id_;_2032870_;_script infofile_;_ZIP::ssf7.xml_;_
    Window("Orders").Dialog("New order").WinEdit("txt_City:").Set DataTable("sCity", dtGlobalSheet) @@ hightlight id_;_590700_;_script infofile_;_ZIP::ssf8.xml_;_
    Window("Orders").Dialog("New order").WinEdit("State:").Set DataTable("sState", dtGlobalSheet) @@ hightlight id_;_460842_;_script infofile_;_ZIP::ssf9.xml_;_
    Window("Orders").Dialog("New order").WinEdit("txt_Zip:").Set DataTable("sZip", dtGlobalSheet) @@ hightlight id_;_1181522_;_script infofile_;_ZIP::ssf10.xml_;_

    'Update the text property to the datatable value to select the radio button
    Window("Orders").Dialog("New order").WinRadioButton("rbn_Card").SetTOProperty "Text", DataTable("sCard",dtGlobalSheet) 
    Window("Orders").Dialog("New order").WinRadioButton("rbn_Card").Set

    'Type keys to set the ExpirationDate 
    Window("Orders").Dialog("New order").WinCalendar("ExpirationDate").Click
    Window("Orders").Dialog("New order").WinCalendar("ExpirationDate").Type(d_ExpiryDate)
   
    Window("Orders").Dialog("New order").WinButton("btn_OK").Click
    
    'Verify the orders list view is created
    If NOT(Window("Orders").WinListView("lvw_Orders").GetItemsCount()=i) Then
           Reporter.ReportEvent micFail,"Orders List View","Orders List View is not created"
    End If

Next



'Function call to fetch Perticular row value
FuncRowValue() @@ hightlight id_;_1705036_;_script infofile_;_ZIP::ssf33.xml_;_


 

  
  


 @@ hightlight id_;_2820464_;_script infofile_;_ZIP::ssf29.xml_;_
 @@ hightlight id_;_67986_;_script infofile_;_ZIP::ssf30.xml_;_

