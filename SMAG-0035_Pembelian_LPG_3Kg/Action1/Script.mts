Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_IDPangkalan, dt_IDAgen, dt_JmlTabung, dt_PinTrx, dt_SaveName, dt_TglPengiriman
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG0035_Pembelian_LPG_3Kg.xlsx", "SMAG-0035")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))


REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_Lpg3kg()

If dt_TCID = "MAG0035-003" Then
	Call Choice_Fav()
	Call Input_IDAgen_Fav(dt_IDAgen)
	Call Input_JmlTabung_Fav(dt_JmlTabung)
	Call Input_Pengiriman(dt_TglPengiriman)
	Device("Device").App("BNI Agen46").MobileButton("Lanjut").Tap
Else
	Call Input_Baru(dt_TCID, dt_IDPangkalan, dt_IDAgen, dt_JmlTabung, dt_TglPengiriman, dt_SaveName)
End If

If dt_TCID = "MAG0035-006" or dt_TCID = "MAG0035-007" Then
	Call Verify_Fail()
Else
	Call Input_Pass(dt_TCID, dt_PinTrx)
End If

If dt_TCID = "MAG0035-001" or dt_TCID = "MAG0035-003" Then
	Call Verify_Success()
ElseIf dt_TCID = "MAG0035-004" or dt_TCID = "MAG0035-005" Then
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0035-002" Then
	Call Verify_Success_Save()
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_LPG.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_LPG.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

    Rem --------- LPG
   	dt_IDPangkalan  	= DataTable.Value("ID_PANGKALAN", dtLocalSheet)
	dt_IDAgen    		= DataTable.Value("ID_AGEN", dtLocalSheet)
	dt_JmlTabung		= DataTable.Value("JMLH_TABUNG", dtLocalSheet)
    	dt_PinTrx       		= DataTable.Value("PIN_TRX", dtLocalSheet)
    	dt_SaveName  	= DataTable.Value("SAVE_NAME", dtLocalSheet)
    	dt_TglPengiriman	= DataTable.Value("TGL_PENGIRIMAN", dtLocalSheet)
    	
    		Rem --------- PLN Manual Advice
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
