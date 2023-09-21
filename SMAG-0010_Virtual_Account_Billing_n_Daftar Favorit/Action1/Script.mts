Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password, dt_NoVA, dt_Nominal, dt_NamaSaveFav, dt_PasswordTrx

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0010_Virtual_Account_Billing_n_Daftar Favorit.xlsx", "SMAG-0010")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))


'REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call GoTo_VirtualAccount()
Call New_Add_VA(dt_TCID, dt_NoVA, dt_Nominal, dt_NamaSaveFav)

'If dt_TCID <> "MAG0010-010" or dt_TCID <> "MAG0010-012" Then
'	Call Input_Pass_Trx(dt_TCID, dt_PasswordTrx, dt_Nominal)
'End If

If dt_TCID = "MAG0010-001" or dt_TCID = "MAG0010-002" or  dt_TCID = "MAG0010-003" or dt_TCID = "MAG0010-004" or dt_TCID = "MAG0010-005" Then
	Call Verif_Success()
ElseIf dt_TCID = "MAG0010-006" Then
	Call Verify_Success_Save()
ElseIf dt_TCID = "MAG0010-007" Then
	Call Verif_Success()
ElseIf dt_TCID = "MAG0010-008" Then
	Call Verif_Fail_Tf_VA()
ElseIf dt_TCID = "MAG0010-010" Then
	Call Verif_Fail_Wrong_VA()
ElseIf dt_TCID = "MAG0010-011" Then
	Call Verif_SaldoKurang_VA()
ElseIf dt_TCID = "MAG0010-012" Then
	Call Verif_Fail_Dormant()
End If

'REM ------ Report Save
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_VirtualAccount.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_VirtualAccount.tsr")
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
	dt_NoVa				= DataTable.Value("NO_VA", dtLocalSheet)
	dt_Nominal			= DataTable.Value("NOMINAL", dtLocalSheet)
	dt_NamaSaveFav		= DataTable.Value("NAMA_SAVE_FAV", dtLocalSheet)
	dt_PasswordTrx		= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
End Sub
