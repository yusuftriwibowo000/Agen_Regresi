Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_NoTelpon, dt_NamaSaveFav, dt_PasswordTrx
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0013_Pulsa_Prabayar_n_Daftar_Favorit.xlsx", "SMAG-0013")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call GoTo_TagihanTelpon()

If dt_TCID = "MAG0013-006" Then
	Call Choice_Fav()
Else 
	Call Input_Baru(dt_TCID, dt_NoTelpon, dt_NamaSaveFav)
End If

Call Input_Pass(dt_TCID, dt_PasswordTrx)

If dt_TCID = "MAG0013-001" or dt_TCID = "MAG0013-002" or dt_TCID = "MAG0013-003" or dt_TCID = "MAG0013-004" or dt_TCID = "MAG0013-006" or dt_TCID = "MAG0013-010" Then
	If dt_TCID = "MAG0013-010" Then
		Call Verify_Success_Dormant()
	Else
		Call Verify_Success()
	End If
ElseIf dt_TCID = "MAG0013-005" Then
	Call Verify_Success_Save()
ElseIf dt_TCID = "MAG0013-007" or dt_TCID = "MAG0013-008" or dt_TCID = "MAG0013-009" or dt_TCID = "MAG0013-011" or dt_TCID = "MAG0013-012" Then
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
	
	REM ---- Tagihan Listrik Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_PulsaPrabayar.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_PulsaPrabayar.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pulsa Prabayar
	dt_NoTelpon			= DataTable.Value("NO_HP", dtLocalSheet)
	dt_NamaSaveFav		= DataTable.Value("NAME_SAVE", dtLocalSheet)
	dt_PasswordTrx		= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
	dt_Username			= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password			= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
