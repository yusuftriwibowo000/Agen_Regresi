Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password, dt_PasswordBaru, dt_Email, dt_Token, dt_PasswordBaruTrx, dt_PasswordBaruTrxNew

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0015_Keamanan_Akun.xlsx", "SMAG-0015")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Scenario Name : " & "Keamanan Akun"))
Call Login(dt_Username, dt_Password)
If dt_TCID = "MAG0015-001" or dt_TCID = "MAG0015-002" or dt_TCID = "MAG0015-006" or dt_TCID = "MAG0015-007"  Then
	Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Username : " & dt_Username, "Password : " & dt_Password, "Password Baru : " & dt_PasswordBaru))
ElseIf dt_TCID = "MAG0015-008" or dt_TCID = "MAG0015-009"  Then
	Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Username : " & dt_Username, "Password : " & dt_Password, "Password Transaksi : " & dt_PasswordBaruTrx, "Password Transaksi Baru : " & dt_PasswordBaruTrxNew))
ElseIf dt_TCID = "MAG0015-003" Then
	Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Username : " & dt_Username, "Password : " & dt_Password, "Password Baru : " & dt_PasswordBaru, "Email : " & dt_Email))
ElseIf dt_TCID = "MAG0015-004" Then
	Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Username : " & dt_Username, "Password : " & dt_Password, "Password Transaksi Baru : " & dt_PasswordBaruTrxNew, "Token : " & dt_Token))
End  If

REM ------- KeagenanMobile
'If dt_TCID = "MAG0015-001" or  dt_TCID = "MAG0015-006" Then
'	Call Pass_Login()
'End  If
If dt_TCID = "MAG0015-001" or dt_TCID = "MAG0015-006" or dt_TCID = "MAG0015-007" Then
	Call Ubah_Pass_Login(dt_Password, dt_PasswordBaru, dt_TCID)
ElseIf dt_TCID = "MAG0015-002" or dt_TCID = "MAG0015-008" or dt_TCID = "MAG0015-009" Then
	Call Ubah_Pass_TRX()
ElseIf dt_TCID = "MAG0015-003" Then
	Call Lupa_Pass_TRX()
ElseIf dt_TCID = "MAG0015-004" Then
	Call Veriv_Lupa_Pass_TRX()
End If
'Call Biometric_Setup_Login()
'Call Input_PassTrx(dt_PasswordBaruTrx)
'Call Pass_Verify_Biometric()
'If dt_TCID = "MAG0015-004" or dt_TCID = "MAG0015-009"  Then
	Call Logout()
'End  If

REM ------ Report Save
Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathKeagenanMobile, LibReport, LibRepo, objSysInfo
	Dim tempKeagenanMobilePath, tempKeagenanMobilePath2, PathKeagenanMobile
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempKeagenanMobilePath 	= Environment.Value("TestDir")
	tempKeagenanMobilePath2 = InStrRev(tempKeagenanMobilePath, "\")
	PathKeagenanMobile 		= Left(tempKeagenanMobilePath, tempKeagenanMobilePath2)
	
	LibPathKeagenanMobile		= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Lib_Keagenan_Mobile\"
	LibReport					= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\LibReport\"
	LibRepo						= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Repo_Keagenan_Mobile\"

	REM ------- Report Library @@ hightlight id_;_525514_;_script infofile_;_ZIP::ssf1.xml_;_
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_Keagenan_Mobile.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Keagenan_Mobile.tsr")
	
	REM ---- Login Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Login.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Login.tsr")
	
	REM ---- Keamanan Akun Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_KeamananAkun.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_KeamananAkun.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Login
	dt_Username			= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password			= DataTable.Value("PASSWORD", dtLocalSheet) @@ hightlight id_;_526150_;_script infofile_;_ZIP::ssf2.xml_;_
	dt_PasswordBaru		= DataTable.Value("PASSWORD_BARU", dtLocalSheet)
	dt_Email				= DataTable.Value("EMAIL", dtLocalSheet)
	dt_Token				= DataTable.Value("TOKEN", dtLocalSheet)
	dt_PasswordBaruTrx	= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
	dt_PasswordBaruTrxNew	= DataTable.Value("PASSWORD_TRX_BARU", dtLocalSheet)
End Sub
