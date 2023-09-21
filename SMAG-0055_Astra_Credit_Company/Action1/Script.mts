REM --- Ncs Multi Finance
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_NoKontrak, dt_PinTrx, dt_SaveName
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0055_Astra_Credit_Company.xlsx", "SMAG-0055")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))

REM ------- Keagenan Mobile
'Call Isi_field_Login(dt_Username, dt_Password)
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_MultiFinance()

If dt_TCID = "MAG0055-003" Then 
	Call Choice_Fav()
Else
	Call Input_Baru_ACC(dt_TCID, dt_NoKontrak, dt_SaveName)
End If

If dt_TCID = "MAG0055-007" Then
	Call Verify_Fail()
Else
	Call Input_Pass(dt_TCID, dt_PinTrx)
End If

If dt_TCID = "MAG0055-001" or dt_TCID = "MAG0055-003" Then
	Call Verify_Success()
ElseIf dt_TCID = "MAG0055-002" Then
	Call Verify_Success_Save()
ElseIf dt_TCID = "MAG0055-004" Then
	Call Verify_Fail_Password_Kosong()
ElseIf dt_TCID = "MAG0055-005" or dt_TCID = "MAG0055-009" Then
	Call Verify_Fail()
End If

REM ------ Report Save
Call spReportSave()

REM ------- Load Function & Repository
Sub spLoadLibrary()
	Dim LibPathKeagenanMobile, LibReport, LibRepo, objSysInfo
	Dim tempKeagenanMobilePath, tempKeagenanMobilePath2, PathKeagenanMobile
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempKeagenanMobilePath 	= Environment.Value("TestDir")
	tempKeagenanMobilePath2 	= InStrRev(tempKeagenanMobilePath, "\")
	PathKeagenanMobile 		= Left(tempKeagenanMobilePath, tempKeagenanMobilePath2)
	
	LibPathKeagenanMobile	= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Lib_Keagenan_Mobile\"
	LibReport					= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\LibReport\"
	LibRepo					= PathKeagenanMobile & "Lib_Repo_Excel_Keagenan_Mobile\Repo_Keagenan_Mobile\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_Keagenan_Mobile.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Keagenan_Mobile.tsr")
	
	REM ---- Login Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Login.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Login.tsr")
	
	REM ---- Personal Loan Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_NCS.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_NCS.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting Keagenan Mobile
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- NCS Multifinance
	dt_NoKontrak	= DataTable.Value("NO_KONTRAK", dtLocalSheet)
	dt_PinTrx	    = DataTable.Value("PIN_TRX", dtLocalSheet)
   	dt_SaveName	    = DataTable.Value("SAVE_NAME", dtLocalSheet)
   	
   	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
