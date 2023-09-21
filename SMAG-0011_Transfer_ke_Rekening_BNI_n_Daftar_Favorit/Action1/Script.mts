Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password, dt_RekeningTujuan, dt_Nominal, dt_NamaSaveFav, dt_PasswordTrx

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0011_Transfer_ke_Rekening_BNI_n_Daftar_Favorit.xlsx", "SMAG-0011")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call GoTo_TransferBNI()

If dt_TCID = "MAG0011-003" Then
	Call Choice_Fav()
	Call Input_Nominal(dt_Nominal)
Else
	Call New_Add_Fav_List(dt_TCID, dt_RekeningTujuan, dt_Nominal, dt_NamaSaveFav)
End If
'Call Input_Pass_Trx(dt_PasswordTrx, dt_TCID)


If dt_TCID = "MAG0011-001" Then
	Call Verif_Trx()
ElseIf dt_TCID = "MAG0011-002" Then
	Call Verif_Trx_Save()
ElseIf dt_TCID = "MAG0011-003" Then
	Call Verif_Trx()
ElseIf dt_TCID = "MAG0011-004" Then
	Call Verif_Wrong_Pin_Trx()
ElseIf dt_TCID = "MAG0011-005" Then
	Call Verif_Max_Limit_Trx()
ElseIf dt_TCID = "MAG0011-006" Then
	Call Verif_Saldo_Kurang_Trx()
ElseIf dt_TCID = "MAG0011-007" Then
	Call Verif_Rekening_Tutup_Trx()
ElseIf dt_TCID = "MAG0011-008" Then
	Call Verif_Rekening_DebtDormant_Trx()
ElseIf dt_TCID = "MAG0011-009" Then
	Call Verif_Rekening_NotFound_Trx()
End If

Call Logout()

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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_TransferRekeningBNI.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_TransferRekeningBNI.tsr")
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
	dt_RekeningTujuan	= DataTable.Value("REKENING_TUJUAN", dtLocalSheet)
	dt_Nominal			= DataTable.Value("NOMINAL", dtLocalSheet)
	dt_NamaSaveFav		= DataTable.Value("NAMA_SAVE_FAV", dtLocalSheet)
	dt_PasswordTrx		= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
End Sub
