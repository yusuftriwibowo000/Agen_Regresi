REM -- I Saku
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_NoHp, dt_PinTrx, dt_SaveName
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0021_TopUp_iSaku_n_Daftar_Favorite.xlsx", "SMAG-0021")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & "Top Up iSaku & Daftar Favorit"))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
'Call Open_MAgen()
Call Goto_ToUp_iSaku()
If dt_TCID <> "MAG0021-003" Then
	Call Input_Baru(dt_TCID, dt_NoHp)
End If
If dt_TCID = "MAG0021-001" Then
	Call Pilih_Nominal(dt_TCID, dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0021-002" Then
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success_Save()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0021-003" Then
	Call Choice_Fav()
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success()
ElseIf dt_TCID = "MAG0021-004" or dt_TCID = "MAG0021-005" or dt_TCID = "MAG0021-007" or dt_TCID = "MAG0021-008"Then
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Fail()
'	Call Back_Wrong_Pin()
ElseIf dt_TCID = "MAG0021-006" Then
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
	
	REM ---- History Transaksi Keagenan Mobile Library

LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_TopUpiSaku.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_TopUpiSaku.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	Rem --------- Paket Data
	dt_NoHp			= DataTable.Value("NO_HP", dtLocalSheet)
	dt_PinTrx			= DataTable.Value("PIN_TRX", dtLocalSheet)
	dt_SaveName			= DataTable.Value("SAVE_NAME", dtLocalSheet)
	
	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
