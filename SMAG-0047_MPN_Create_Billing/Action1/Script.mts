REM --- Pajak Daerah
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_KodeBayar, dt_PinTrx, dt_SaveName
Dim dt_Username, dt_Password
Dim dt_NPWP, dt_Bulan, dt_Tahun,  dt_NPWP_Penyetor, dt_Nominal

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0047_MPN_Create_Billing.xlsx", "SMAG-0047")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))

Call Input_Pembuatan_Kode_Biling(dt_NPWP, dt_Bulan, dt_Tahun, dt_NPWP_Penyetor, dt_Nominal)

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_PajakMPN()
Call Input_Pembuatan_Kode_Biling(dt_NPWP, dt_Bulan, dt_Tahun, dt_NPWP_Penyetor, dt_Nominal)

Call Input_Pass(dt_TCID, dt_PinTrx)
Call Verify_Success_Save()
Call Verification()

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
	
	REM ---- Personal Loan Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_MPN_CreateBilling.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_MPN_Createbilling.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting Keagenan Mobile
	dt_TCID				= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc	= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc	= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult	= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pajak Daerah
	dt_PinTrx	    		= DataTable.Value("PIN_TRX", dtLocalSheet)
  	
  	   	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
	
'	Pembuatan Kode Billing
	dt_NPWP			= DataTable.Value("NO_NPWP", dtLocalSheet)
	dt_Bulan			= DataTable.Value("MASA_KERJA_BULAN", dtLocalSheet)
	dt_Tahun			= DataTable.Value("MASA_KERJA_TAHUN", dtLocalSheet)
	dt_NPWP_Penyetor	= DataTable.Value("NO_NPWP_PENYETOR", dtLocalSheet)
	dt_Nominal			= DataTable.Value("NOMINAL", dtLocalSheet)
	
End Sub
