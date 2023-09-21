REM -- ZIS Qurban Automation test
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Nominal, dt_NoHp, dt_PinTrx
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0043_ZIS_Qurban.xlsx", "SMAG-0043")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))



REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_ZISQurban()
Call Input_Baru(dt_TCID, dt_Nominal)

If dt_TCID = "MAG0043-008" or dt_TCID = "MAG0043-009" Then
	Call Verify_Fail()
Else
	Call Input_Pass(dt_TCID, dt_PinTrx)
End If

If dt_TCID = "MAG0043-001" or dt_TCID = "MAG0043-002" or dt_TCID = "MAG0043-003" or dt_TCID = "MAG0043-004" Then
	Call Verify_Success()
Else
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
	
	REM ---- BPJSK Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_ZISQurban.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_ZISQurban.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pulsa Prabayar
	dt_Nominal	        = DataTable.Value("NOMINAL", dtLocalSheet)
	dt_NoHp		        = DataTable.Value("NO_HP", dtLocalSheet)
	dt_PinTrx	    	= DataTable.Value("PIN_TRX", dtLocalSheet)
	
	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub

