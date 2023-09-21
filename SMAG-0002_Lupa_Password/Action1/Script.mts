Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim  dt_Username, dt_Password, dt_PasswordBaru, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir


REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0002_Lupa_Password.xlsx", "SMAG-0002")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))


REM ------- Keagenan Mobile Andorid
Call Go_To_Lupa_Pass()
If dt_TCID = "MAG0002-001" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verif_Link_Reset_Was_Sent()
	Call Open_Email()
	Call Verify_Email_was_Open()
	Call Veriv_Change_Pass(dt_PasswordBaru, dt_TCID)
	Call Verif_Otp_Was_Sent()
	Call Input_OTP_Manual()
	Call Verify_Success()
	Call Pass_Login(dt_Username, dt_Password)
ElseIf dt_TCID = "MAG0002-002" or dt_TCID = "MAG0002-003" or dt_TCID = "MAG0002-008" or dt_TCID = "MAG0002-009" or dt_TCID = "MAG0002-010" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0002-011" or dt_TCID = "MAG0002-012" or dt_TCID = "MAG0002-013" or dt_TCID = "MAG0002-014" or dt_TCID = "MAG0002-015" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verify_Fail_Empty()
ElseIf dt_TCID = "MAG0002-004" or dt_TCID = "MAG0002-016" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verif_Link_Reset_Was_Sent()
	Call Open_Email()
	Call Verify_Email_was_Open()
	Call Veriv_Change_Pass(dt_PasswordBaru, dt_TCID)
ElseIf dt_TCID = "MAG0002-005" or dt_TCID = "MAG0002-006" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verif_Link_Reset_Was_Sent()
	Call Open_Email()
	Call Verify_Email_was_Open()
	Call Veriv_Change_Pass(dt_PasswordBaru, dt_TCID)
	Call Verif_Otp_Was_Sent()
	Call Verify_OTP(dt_TCID)
ElseIf dt_TCID = "MAG0002-017" Then
	Call Isi_Form_VerifikasiData(dt_TCID, dt_Username, dt_Email, dt_NoIdentitas, dt_NoRekening, dt_TglLahir)
	Call Verif_Link_Reset_Was_Sent()
	Call Open_Email()
	Call Verify_Email_was_Open()
	Call Veriv_Change_Pass(dt_PasswordBaru, dt_TCID)
	Call Verif_Otp_Was_Sent()
	Call Resend_OTP()
End If

REM ------ Report Save
'Call spReportSave()
	
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_LupaPassword.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_LupaPassword.tsr")
	Call RepositoriesCollection.Add(LibRepo & "repo_gabungan _login.tsr")
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
	dt_PasswordBaru		= DataTable.Value("PASSWORD_BARU", dtLocalSheet)
	dt_Email				= DataTable.Value("EMAIL", dtLocalSheet)
	dt_NoIdentitas		= DataTable.Value("NO_IDENTITAS", dtLocalSheet)
	dt_NoRekening		= DataTable.Value("NO_REKENING", dtLocalSheet)
	dt_TglLahir				= DataTable.Value("TGL_LAHIR", dtLocalSheet)
End  Sub


