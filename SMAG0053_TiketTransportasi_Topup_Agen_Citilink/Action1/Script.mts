REM -- Tiket Transportasi Penerbangan
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_PayCode, dt_PinTrx
Dim dt_Username, dt_Password, dt_Rekening, dt_SaveName

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG0053_TiketTransportasi_Topup_Agen_Citilink.xlsx", "SMAG-0053")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & "BAYAR & BELI Tiket Transportasi Penerbangan & Daftar Favorite"))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_Topup_Citilink()
Call New_Add_Fav_List_Topup_Citilink(dt_TCID, dt_Rekening, dt_Nominal, dt_SaveName)

If dt_TCID = "MAG0053-001" or dt_TCID = "MAG0053-003" Then
	Call Input_Pass_Trx(dt_TCID, dt_PinTrx)
	Call Verify_Success()
ElseIf dt_TCID = "MAG0053-002" Then
	Call Input_Pass_Trx(dt_TCID, dt_PinTrx)
	Call Verify_Success_Save()
ElseIf dt_TCID = "MAG0053-004" or dt_TCID = "MAG0053-005" or dt_TCID = "MAG0053-006" or dt_TCID = "MAG0053-008" Then 
	Call Input_Pass_Trx(dt_TCID, dt_PinTrx)
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0053-007" Then
	Call Verify_Fail()
End If
Call Logout()

'Call Input_PayCode(dt_TCID, dt_PayCode)
''print(dt_TCID)
''If dt_TCID <> "MAG0026-007" or dt_TCID <> "MAG0026-008" Then
''	Call Input_Pass(dt_TCID, dt_PinTrx)
''End If
'If dt_TCID = "MAG0026-001" or dt_TCID = "MAG0026-002" or dt_TCID = "MAG0026-003" or dt_TCID = "MAG0026-004" Then
'	Call Verify_Success()
'ElseIf dt_TCID = "MAG0026-005" or dt_TCID = "MAG0026-006" or dt_TCID = "MAG0026-007" or dt_TCID = "MAG0026-008" or dt_TCID = "MAG0026-009" Then
'	Call Verify_Fail()
'End If
'
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_TiketTransportasi.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_TiketTransportasi.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	Rem --------- Paket Data
	dt_PayCode			= DataTable.Value("KODE_BAYAR", dtLocalSheet)
	dt_PinTrx			= DataTable.Value("PIN_TRX", dtLocalSheet)
	dt_Rekening		= DataTable.Value("KODE_BAYAR", dtLocalSheet)
	dt_SaveName	= DataTable.Value("NAMA_SAVE", dtLocalSheet)
	
	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub


