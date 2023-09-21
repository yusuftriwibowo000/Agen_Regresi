Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0016_Riwayat_Pesan.xlsx", "SMAG-0016")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & "Riwayat Pesan"))

'Call Pilih_VerifyHapusPesan("MAG0016-004")


REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
'Call Open_MAgen()
Call Go_toAkun()
Call Go_toRiwayatPesan()
If dt_TCID = "MAG0016-001" Then
	Call Go_toBelumTerbaca()
	Call Go_toBroadCastMessage()
	Call Pilih_VerifyOpenPesan(dt_TCID)
	Call CameBack()
	Call Logout()
ElseIf dt_TCID = "MAG0016-002" Then
	Call Go_toBelumTerbaca()
	Call Go_toBroadCastMessage()
	Call Pilih_VerifyHapusPesan(dt_TCID)
	Call CameBack2()
	Call Logout()
ElseIf dt_TCID = "MAG0016-003" Then
	Call Go_toPesanTerbaca()
	Call Go_toBroadCastMessage()
	Call Pilih_VerifyOpenPesan(dt_TCID)
	Call CameBack()
	Call Logout()
ElseIf dt_TCID = "MAG0016-004" Then
	Call Go_toPesanTerbaca()
	Call Go_toBroadCastMessage()
	Call Pilih_VerifyHapusPesan(dt_TCID)
	Call CameBack2()
	Call Logout()
ElseIf dt_TCID = "MAG0016-005" Then
	Call Go_toPesanTerhapus()
	Call Go_toBroadCastMessage()
	Call Pilih_VerifyOpenPesan(dt_TCID)
	Call CameBack()
	Call Logout()
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_RiwayatPesan.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_RiwayatPesan.tsr")
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
	
	Rem --------- Riwayat Pesan
End Sub
