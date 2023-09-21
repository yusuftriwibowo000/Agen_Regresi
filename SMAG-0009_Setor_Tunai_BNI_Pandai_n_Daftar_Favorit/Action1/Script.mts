Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password, dt_Rekening, dt_Nominal, dt_NamaSaveFav, dt_PasswordTrx

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0009_Setor_Tunai_Pandai.xlsx", "SMAG-0009")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call GoTo_LayananKeuangan()
Call New_Add_Fav_List(dt_TCID, dt_Rekening, dt_Nominal, dt_NamaSaveFav)
If dt_TCID = "MAG0009-001" or dt_TCID = "MAG0009-003" Then
	Call Input_Pass_Trx(dt_TCID, dt_PasswordTrx)
	Call Verify_Success()
ElseIf dt_TCID = "MAG0009-002" Then
	Call Input_Pass_Trx(dt_TCID, dt_PasswordTrx)
	Call Verify_Success_Save()
ElseIf dt_TCID = "MAG0009-004" or dt_TCID = "MAG0009-007" or dt_TCID = "MAG0009-008" Then 
	Call Input_Pass_Trx(dt_TCID, dt_PasswordTrx)
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0009-006" Then
	Call Verify_Fail()
End If
Call Logout()

REM ------ Report Save
Call spReportSave()
	
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
	
	REM ---- History Transaksi Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_SetorTunaiPandai.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_SetorTunaiPandai.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Login
	dt_Username			= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password			= DataTable.Value("PASSWORD", dtLocalSheet)
	dt_Rekening			= DataTable.Value("NO_REKENING", dtLocalSheet)
	dt_Nominal			= DataTable.Value("NOMINAL", dtLocalSheet)
	dt_NamaSaveFav		= DataTable.Value("NAMA_SAVE_FAV", dtLocalSheet)
	dt_PasswordTrx		= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
End Sub
