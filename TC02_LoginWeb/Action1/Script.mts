
Desktop.RunAnalog "Track1"
'Username = trim((DataTable("Username","Global")))
'Password = trim((DataTable("Password","Global")))
'Const Path_Data ="D:\TestUFT\Excel"
'Const Data ="TestWeb.xlsx"

DataTable.Import "F:\TrainingUFT\Excel\testFlight.xlsx"
iRowCount = DataTable.getSheet("TC02_LoginWeb").GetRowCount
For i = 1 To iRowCount 
    DataTable.SetCurrentRow(i)
	Username = trim((DataTable.Value("Username","TC02_LoginWeb")))
	Password = trim((DataTable.Value("Password","TC02_LoginWeb")))
	
	WebClickLink "Advantage Shopping","Advantage Shopping","UserMenu"
	wait 2
	WebBrowserEdit "Advantage Shopping","Advantage Shopping","username",Username @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu")_;_script infofile_;_ZIP::ssf1.xml_;_
	wait 2
	WebBrowserEdit "Advantage Shopping","Advantage Shopping","password",Password @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("username")_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 2
	WebClickButton "Advantage Shopping","Advantage Shopping","sign_in_btnundefined" @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("password")_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 2
	WebClickLink "Advantage Shopping","Advantage Shopping","UserMenu_2"
	wait 2
	WebClickLink "Advantage Shopping","Advantage Shopping","Link"
	wait 2
	
Next

i = Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("dvantage").GetROProperty("innertext")
print i

