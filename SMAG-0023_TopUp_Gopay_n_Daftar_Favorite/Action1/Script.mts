Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_NoHp, dt_PinTrx, dt_SaveName, dt_Nominal
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0023_TopUp_Gopay_n_Daftar_Favorite.xlsx", "SMAG-0023")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & "Top Up Gopay & Daftar Favorite"))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_ToUp_Gopay(dt_TCID)
If dt_TCID <> "MAG0023-005" Then
	Call Input_Baru(dt_TCID, dt_NoHp)
End If
If dt_TCID = "MAG0023-001" or dt_TCID = "MAG0023-002" or dt_TCID = "MAG0023-003" Then
	Call Pilih_Nominal(dt_TCID, dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0023-004" Then
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success_Save()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0023-005" Then
	Call Choice_Fav()
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Success()
ElseIf dt_TCID = "MAG0023-006"  or dt_TCID = "MAG0023-009" or dt_TCID = "MAG0023-010" Then
	Call Pilih_Nominal(dt_TCID,dt_SaveName)
	Call Input_Pass(dt_TCID, dt_PinTrx)
	Call Verify_Fail()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0023-008" Then
	Call Verify_Fail()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0023-011" or dt_TCID = "MAG0023-012" Then
	Call Nominal_Lainnya(dt_Nominal)
	Call Verify_Block_Button()
'	Call Back_Wrong()
ElseIf dt_TCID = "MAG0023-007" Then
Call Nominal_Lainnya(dt_Nominal)
Call Input_Pass(dt_TCID, dt_PinTrx)
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_TopUpGopay.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_TopUpGopay.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	Rem --------- Paket Data
	dt_NoHp			= DataTable.Value("NO_HP", dtLocalSheet)
	dt_PinTrx		= DataTable.Value("PIN_TRX", dtLocalSheet)
	dt_SaveName		= DataTable.Value("SAVE_NAME", dtLocalSheet)
	dt_Nominal		= DataTable.Value("NOMINAL", dtLocalSheet)
	
	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub


