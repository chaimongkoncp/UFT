'Open App
OpenAppGUI "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe"
'Improt
DataTable.Import "F:\TrainingUFT\Excel\testFlight.xlsx"
'Get Row
rowcount = DataTable.GetSheet("TC01_Login").GetRowCount
For i = 1 To DataTable.GetSheet("TC01_Login").GetRowCount
			 DataTable.LocalSheet.SetCurrentRow(i)
			 Username = trim((DataTable("Username","TC01_Login")))
			 Password = trim((DataTable("Password","TC01_Login")))
			 Name = trim((DataTable("Name","TC01_Login")))
			 FromCity = trim((DataTable("FromCity","TC01_Login")))
			 ToCity = trim((DataTable("ToCity","TC01_Login")))
			 
If i = 1 Then
	InputText "Micro Focus MyFlight Sample","agentName",Username
	InputText "Micro Focus MyFlight Sample","password",Password
	ClickButton "Micro Focus MyFlight Sample","OK" @@ hightlight id_;_2132908840_;_script infofile_;_ZIP::ssf27.xml_;_
	
End If @@ hightlight id_;_1918280624_;_script infofile_;_ZIP::ssf18.xml_;_

WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1918282448_;_script infofile_;_ZIP::ssf20.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,0 @@ hightlight id_;_1918283840_;_script infofile_;_ZIP::ssf21.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_1918284080_;_script infofile_;_ZIP::ssf22.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set Name @@ hightlight id_;_2094598496_;_script infofile_;_ZIP::ssf24.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click

order = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 87 completed").GetROProperty("text")

DataTable.Value("Order","TC01_Login") = order
WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click

Next
 @@ hightlight id_;_2126421496_;_script infofile_;_ZIP::ssf28.xml_;_
WpfWindow("Micro Focus MyFlight Sample").Close @@ hightlight id_;_7277036_;_script infofile_;_ZIP::ssf26.xml_;_
DataTable.Export "D:\TestUFT\Excel\testOrder.xlsx"
