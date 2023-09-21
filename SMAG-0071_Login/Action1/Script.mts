Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password
Dim dtVerifUsername, dtVerifEmail, dtVerifNoIdentitas, dtVerifNoRek, dtTglLahir


REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0071_Login.xlsx", "Sheet1")
Call spGetDatatable()
	

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))


Call Login(dt_Username, dt_Password)
REM -------------- Keagenan Mobile Android Login
If dt_TCID = "MAG0001-001" Then
	Call Blocked_User(dt_Username, dt_Password)
End If
If dt_TCID = "MAG0071-011" or dt_TCID = "MAG0071-012" or dt_TCID = "MAG0071-013" or dt_TCID = "MAG0071-014" or dt_TCID = "MAG0071-015" or dt_TCID = "MAG0071-016" or dt_TCID = "MAG0071-017" or dt_TCID = "MAG0071-028" or dt_TCID = "MAG0071-029" or dt_TCID = "MAG0071-030" or dt_TCID = "MAG0071-031"Then
	Call Lupa_Password()
	else
	Call Blocked_User(dt_Username, dt_Password)
End If
Call Input_Verifikasi_Data(dt_TCID, dtVerifUsername, dtVerifEmail, dtVerifNoIdentitas, dtVerifNoRek, dtTglLahir)
If dt_TCID = "MAG0071-015" or dt_TCID = "MAG0071-016" or dt_TCID = "MAG0071-017" or dt_TCID = "MAG0071-028" or dt_TCID = "MAG0071-029" or dt_TCID = "MAG0071-030" or dt_TCID = "MAG0071-031" Then
	Call Password_Login_Baru(dt_Password)
End If
Call Verifikasi_Data(dt_TCID)
REM -------------- Report Save
Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathKeagenanMobile, LibReport, LibRepo, objSysInfo
	Dim tempKeagenanMobilePath, tempKeagenanMobilePath2, PathKeagenanMobile
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempKeagenanMobilePath 	= Environment.Value("TestDir")
	tempKeagenanMobilePath2 = InStrRev(tempKeagenanMobilePath, "\")
	PathKeagenanMobile 		= Left(tempKeagenanMobilePath, tempKeagenanMobilePath2)
	
	LibPathKeagenanMobile		= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Lib_Keagenan_Mobile\"
	LibReport						= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\LibReport\"
	LibRepo							= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Repo_Keagenan_Mobile\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_Keagenan_Mobile.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Keagenan_Mobile.tsr")
	
	REM ---- Login Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Login.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "repo_gabungan _login.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
	
	REM ---------InputVerifikasi
	dtVerifUsername	= DataTable.Value("VERIFIKASI_USERNAME", dtLocalSheet)	
	dtVerifEmail			= DataTable.Value("VERIFIKASI_EMAIL", dtLocalSheet)
	dtVerifNoIdentitas	= DataTable.Value("VERIFIKASI_NOMER_IDENTITAS", dtLocalSheet)
	dtVerifNoRek		= DataTable.Value("VERIFIKASI_NOMER_REKENING", dtLocalSheet)
	dtTglLahir			= DataTable.Value("VERIFIKASI_TANGGAL_LAHIR", dtLocalSheet)
End Sub









'AIUtil("text_box", "Link bit .ly/ProgramBukaRekeningAgen46").Type "BNIAG198845"
'AIUtil("text_box", "Password Login").Type "Ipybni06!"
'AIUtil("button", "Masuk").Click
'AIUtil.FindTextBlock("Username").Click


AIUtil.SetContext Device("Device")
AIUtil("text_box", "Username").Type "BNIAG198845"
AIUtil("text_box", "Password Login").Type "Ipybni06!"
AIUtil("button", "Masuk").Click

'Device("Device").App("BNI Agen46_2").MobileObject("Layanan Keuangan").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileObject("setor tunai").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileButton("Buat Input Baru").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Set "1234567892"
'wait 2
'Device("Device").App("BNI Agen46_2").MobileButton("Lanjut").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Set "50000"
'wait 2
'Device("Device").App("BNI Agen46_2").MobileButton("Lanjut").Tap



'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Tap
'wait 1
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 login").Set "testing"
'wait 2
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 input passwrod").Tap
'wait 1
'Device("Device").App("BNI Agen46_2").MobileEdit("testing agen 2.5.0 input passwrod").Set "testing"
'wait 2
'Call CaptureImageUFTV2(Device("Device").App("BNI Agen46"), "isi inputan login", " ", compatibilityMode.Mobile, ReportStatus.Done)
'wait 
'Device("Device").App("BNI Agen46_2").MobileButton("testing agen 2.5.0 button masuk").Tap
'wait 2
'Device("Device").App("BNI Agen46_2").MobileObject("testing agen 2.5.0 button pop up ok").Tap
'
