Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password, dt_TglAwal, dt_TglAkhir

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0014_Laporan_Riwayat_Transaksi.xlsx", "SMAG-0014")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))

'REM ------- Keagenan Mobile Android
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
'
If dt_TCID = "MAG0014-001" Then
	Call Cek_All_Trx()
ElseIf dt_TCID = "MAG0014-002" Then
	Call Cek_Trf_Pulsa_Trx()
ElseIf dt_TCID = "MAG0014-003" Then
	Call Cek_Topup_Ewallet_Trx()
ElseIf dt_TCID = "MAG0014-004" or dt_TCID = "MAG0014-005" Then
	Call Cek_Trx_Period_Minggu_Bulan(dt_TCID)
ElseIf dt_TCID = "MAG0014-006" or dt_TCID ="MAG0014-007" Then 
	Call Cek_Trx_Jangka_Waktu(dt_TCID, dt_TglAwal, dt_TglAkhir)
End If

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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_HistoryTrx.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_HistoryTrx.tsr")
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
	dt_TglAwal	= DataTable.Value("TANGGAL_AWAL", dtLocalSheet)
	dt_TglAkhir = DataTable.Value("TANGGAL_AKHIR", dtLocalSheet)
End Sub
