Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Username, dt_Password
Dim dt_SetoranAwal
Dim dt_NamaDepan, dt_NamaTengah, dt_NamaBelakang, dt_Email, dt_TempatLahir, dt_NoKtp, dt_NPWP
Dim dt_Alamat, dt_RT, dt_RW
Dim dt_NoHp, dt_DescPekerjaan, dt_Penghasilan
Dim dt_NamaPerusahaan, dt_AlamatPerusahaan, dt_KotaPerusahaan, dt_DetaiPekerjaan
Dim dt_PassTrx

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0008_Buka_Rekening_BNI_Pandai.xlsx", "SMAG-0008")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))

REM ------- Keagenan Mobile
'wait 5
Call Login(dt_Username, dt_Password)
If dt_TCID = "MAG0008-001" or dt_TCID = "MAG0008-002" or dt_TCID = "MAG0008-004" Then
	Call Open_MAgen()
	Call Go_Buka_Rekening_Page() 
	Call Input_Setoran_Awal(dt_SetoranAwal)
	Call Konfirmasi_Setoran_Awal()
	Call Input_Data_Diri(dt_NamaDepan, dt_NamaTengah, dt_NamaBelakang, dt_Email, dt_TempatLahir, dt_NoKtp, dt_NPWP)
	wait 4
	Call Act_Lanjutkan()
	wait 4
	Call Input_Alamat(dt_Alamat, dt_RT, dt_RW)
	wait 4
	Call Input_Kontak_Nasabah(dt_NoHp, dt_DescPekerjaan, dt_Penghasilan)
	wait 4
	Call Input_Detail_Pekerjaan(dt_NamaPerusahaan, dt_AlamatPerusahaan, dt_KotaPerusahaan, dt_DetaiPekerjaan)
	wait 4
	Call Upload_n_InputPassTrx(dt_PassTrx)
End If

If dt_TCID = "MAG0008-001" Then
	Call VerifySukses_BukaRekening()
ElseIf dt_TCID = "MAG0008-002" Then
	Call Verify_PasswordTrx_Salah()
ElseIf dt_TCID = "MAG0008-003" Then
	Call Open_MAgen()
	Call Go_Buka_Rekening_Page()
	Call Input_Setoran_Awal(dt_SetoranAwal)
	Call Konfirmasi_Setoran_Awal()
	Call Input_Data_Diri(dt_NamaDepan, dt_NamaTengah, dt_NamaBelakang, dt_Email, dt_TempatLahir, dt_NoKtp, dt_NPWP)
	Call Verify_DataKosong()
ElseIf dt_TCID = "MAG0008-004" Then
	Call Verify_Data_Sudah_Digunakan()
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
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_BukaRekening.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_BukaRekening.tsr")
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
	
	Rem --------- Buka Rekening
	dt_SetoranAwal		= DataTable.Value("SETORAN_AWAL", dtLocalSheet)
	dt_NamaDepan		= DataTable.Value("NAMA_DEPAN", dtLocalSheet)
	dt_NamaTengah		= DataTable.Value("NAMA_TENGAH", dtLocalSheet)
	dt_NamaBelakang		= DataTable.Value("NAMA_BELAKANG", dtLocalSheet)
	dt_Email				= DataTable.Value("EMAIL", dtLocalSheet)
	dt_TempatLahir		= DataTable.Value("TEMPAT_LAHIR", dtLocalSheet)
	dt_NoKtp			= DataTable.Value("NO_KTP", dtLocalSheet)
	dt_NPWP			= DataTable.Value("NPWP", dtLocalSheet)
	dt_Alamat			= DataTable.Value("ALAMAT_LENGKAP", dtLocalSheet)
	dt_RT				= DataTable.Value("RT", dtLocalSheet)
	dt_RW				= DataTable.Value("RW", dtLocalSheet)
	dt_NoHp				= DataTable.Value("NO_HP", dtLocalSheet)
	dt_DescPekerjaan	= DataTable.Value("DESKRIPSI_PEKERJAAN", dtLocalSheet)
	dt_Penghasilan		= DataTable.Value("PENGHASILAN", dtLocalSheet)
	dt_NamaPerusahaan	= DataTable.Value("NAMA_PERUSAHAAN", dtLocalSheet)
	dt_AlamatPerusahaan = DataTable.Value("ALAMAT_PERUSAHAAN", dtLocalSheet)
	dt_KotaPerusahaan = DataTable.Value("KOTA_PERUSAHAAN", dtLocalSheet)
	dt_DetaiPekerjaan = DataTable.Value("DETAIL_PEKERJAAN", dtLocalSheet)
	dt_PassTrx			= DataTable.Value("PASSWORD_TRX", dtLocalSheet)
End Sub
