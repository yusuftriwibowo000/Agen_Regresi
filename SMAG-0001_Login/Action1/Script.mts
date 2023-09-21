Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0001_Login.xlsx", "SMAG-0001")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))
'Call Isi_field_Login("BNIAG198845", "Ipybni06!")
REM -------------- Keagenan Mobile Android Login
If dt_TCID = "MAG0001-001" Then
	Call Pass_Login(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-002" Then
	Call Login_Simpan_Uname(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-003" Then
	Call Wrong_Password(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-004" Then
	Call Empty_Password(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-005" Then
	Call Blocked_User(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-006" Then
	Call Credential_User(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0001-013" or dt_TCID = "MAG0001-014" Then
	Call spLoadingScreenLogin()
	Call Go_To_Login_Biometric(dt_TCID, dt_Username)
	Call verify_passed_login()
End If

If dt_TCID <> "MAG0001-003" and dt_TCID <> "MAG0001-004" and dt_TCID <> "MAG0001-005" and dt_TCID <> "MAG0001-006" Then
	Call Logout()
End If

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
	LibRepo						= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Repo_Keagenan_Mobile\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_Keagenan_Mobile.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Keagenan_Mobile.tsr")
	
	REM ---- Login Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Login.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Login.tsr")
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
End Sub

