Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_KodeRegistrasi, dt_NamaBelakang
Dim dt_NamaLengkap, dt_NoIdentitas, dt_NPWP, dt_Occupation, dt_NoTelpRumah, dt_NoHp, dt_Email
Dim dt_Address, dt_Rt, dt_Rw, dt_Provinsi, dt_City, dt_Kec, dt_Kel, dt_KPos
Dim dt_JenisUsaha, dt_NamaToko
Dim dt_KantorCabang 
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0003_Registrasi.xlsx", "SMAG-0003")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_ScenarioDesc))

REM ------- Keagenan Mobile
If dt_TCID = "MAG0003-001" Then
'	Call Login(dt_Username, dt_Password)
	Call GoTo_Registrasi()
	Call Input_DataDiri(dt_NamaLengkap, dt_NoIdentitas, dt_NPWP, dt_Occupation, dt_NoTelpRumah, dt_NoHp, dt_Email)
	Call Next_Action()
	Call Input_AlamatIdentitas(dt_Address, dt_Rt, dt_Rw, dt_Provinsi, dt_City, dt_Kec, dt_Kel, dt_KPos)
	Call Next_Action()
	Call Input_AlamatUsaha(dt_JenisUsaha, dt_NamaToko)
	Call Next_Action()
	Call Verifikasi_KantorCabang(dt_KantorCabang)
	Call Next_Action()
	Call Kirim()
	Call Verif_Success()
ElseIf dt_TCID = "MAG0003-002" Then
	Call GoTo_Registrasi()
	Call Input_DataDiri(dt_NamaLengkap, dt_NoIdentitas, dt_NPWP, dt_Occupation, dt_NoTelpRumah, dt_NoHp, dt_Email)
	Call Next_Action()
	Call Verif_Salah_NoIdentitas()
ElseIf dt_TCID = "MAG0003-003" Then
	Call GoTo_Registrasi()
	Call Input_DataDiri(dt_NamaLengkap, dt_NoIdentitas, dt_NPWP, dt_Occupation, dt_NoTelpRumah, dt_NoHp, dt_Email)
	Call Next_Action()
	Call Verif_Empty_Field()
ElseIf dt_TCID = "MAG0003-004" Then
	Call GoTo_Registrasi()
	Call Input_DataDiri(dt_NamaLengkap, dt_NoIdentitas, dt_NPWP, dt_Occupation, dt_NoTelpRumah, dt_NoHp, dt_Email)
	Call Next_Action()
	Call Input_AlamatIdentitas(dt_Address, dt_Rt, dt_Rw, dt_Provinsi, dt_City, dt_Kec, dt_Kel, dt_KPos)
	Call Next_Action()
	Call Input_AlamatUsaha(dt_JenisUsaha, dt_NamaToko)
	Call Next_Action()
	Call Verifikasi_KantorCabang(dt_KantorCabang)
	Call Next_Action()
	Call Kirim()
	Call Data_Telah_Digunakan()
ElseIf dt_TCID = "MAG0003-005" Then
	Call Go_To_Pendaftaran()
	Call Cek_Status(dt_NamaBelakang, dt_KodeRegistrasi)
	Call Check_Action()
	Call Cek_Status_Pass()
ElseIf dt_TCID = "MAG0003-006" Then
	Call Go_To_Pendaftaran()
	Call Cek_Status(dt_NamaBelakang, dt_KodeRegistrasi)
	Call Check_Action()
	Call Cek_Status_Fail()
End If

REM ------ Report Save
Call spReportSave()
	
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
	
	REM ---- Registrasi Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Registrasi.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Registrasi.tsr")
	
	REM ---- Login Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_Login.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_Login.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	REM --------- Data Diri Registrasi
	dt_NamaLengkap 		= DataTable.Value("NAMA_LENGKAP", dtLocalSheet)
	dt_NoIdentitas		= DataTable.Value("NOMER_IDENTITAS", dtLocalSheet)
	dt_NPWP				= DataTable.Value("NPWP", dtLocalSheet)
	dt_Occupation		= DataTable.Value("JENIS_PEKERJAAN", dtLocalSheet)		
	dt_NoTelpRumah		= DataTable.Value("NO_TELP_RUMAH", dtLocalSheet)
	dt_NoHp				= DataTable.Value("NO_HP", dtLocalSheet)
	dt_Email			= DataTable.Value("EMAIL", dtLocalSheet)
	
	REM --------- Alamat Identitas Registrasi
	dt_Address	= DataTable.Value("ALAMAT_LENGKAP", dtLocalSheet)
	dt_Rt		= DataTable.Value("RT", dtLocalSheet)
	dt_Rw		= DataTable.Value("RW", dtLocalSheet)
	dt_Provinsi	= DataTable.Value("PROVINSI", dtLocalSheet)
	dt_City		= DataTable.Value("KOTA", dtLocalSheet)
	dt_Kec		= DataTable.Value("KECAMATAN", dtLocalSheet)
	dt_Kel		= DataTable.Value("KELURAHAN", dtLocalSheet)
	dt_KPos		= DataTable.Value("KODE_POS", dtLocalSheet)
	
	REM --------- Alamat Usaha Registrasi
	dt_JenisUsaha	= DataTable.Value("JENIS_USAHA", dtLocalSheet)
	dt_NamaToko		= DataTable.Value("NAMA_TOKO", dtLocalSheet)
	
	REM --------- Verifikasi Vantor Cabang
	dt_KantorCabang	= DataTable.Value("KANTOR_CABANG_PENDANAAN", dtLocalSheet)
	
	Rem --------- Status Registrasi
	dt_NamaBelakang = DataTable.Value("NAMA_BELAKANG", dtLocalSheet)
	dt_KodeRegistrasi = DataTable.Value("KODE_REGISTRASI", dtLocalSheet)
	
	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
	
End Sub
