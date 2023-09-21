Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_NoTelpon, dt_NamaSaveFav, dt_PasswordTrx
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0037_Telpon_or_Selular_Pascabayar.xlsx", "SMAG-0037")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_TagihanTelpon()

If dt_TCID = "MAG0037-006" Then
	Call Choice_Fav()
Else
	Call Input_Baru(dt_TCID, dt_NoTelpon, dt_NamaSaveFav)
End If

If dt_TCID = "MAG0037-009" or dt_TCID = "MAG0037-011" Then
	Call Verify_Fail()
Else
	Call Input_Pass(dt_TCID, dt_PasswordTrx)
End If

If dt_TCID = "MAG0037-007" or dt_TCID = "MAG0037-008" or dt_TCID = "MAG0037-010" Then
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0037-005" Then
	Call Verify_Success_Save()
Else
	Call Verify_Success()
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
	
	REM ---- Telpon atau Selular Pascabayar Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_TelponSelularPascabayar.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_TelponSelularPascabayar.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pulsa Prabayar
	dt_NoTelpon			= DataTable.Value("NO_TELPON", dtLocalSheet)
	dt_NamaSaveFav		= DataTable.Value("SAVE_NAME", dtLocalSheet)
	dt_PasswordTrx		= DataTable.Value("PIN_TRX", dtLocalSheet)
	
	    REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
