﻿REM --- Pajak Daerah
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_KodeBayar, dt_PinTrx, dt_SaveName
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0046_Pajak_Daerah_BPHITB.xlsx", "SMAG-0046")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))

REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen()
Call Goto_BPHITB()

If dt_TCID = "MAG0046-005" Then 
	Call Choice_Fav()
Else
	Call Input_Baru(dt_TCID, dt_KodeBayar, dt_SaveName)
End If

Call Input_Pass(dt_TCID, dt_PinTrx)


If dt_TCID <> "MAG0046-007" Then
	If dt_TCID = "MAG0046-004" Then
		Call Verify_Success_Save()
	Else
		Call Verify_Verification()
	End If
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
	
	REM ---- Personal Loan Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_PajakDaerah.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_PajakDaerah.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting Keagenan Mobile
	dt_TCID						    = DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc		= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pajak Daerah
	dt_KodeBayar		= DataTable.Value("KODE_PEMBAYARAN", dtLocalSheet)
	dt_PinTrx	    	= DataTable.Value("PIN_TRX", dtLocalSheet)
  	dt_SaveName	    = DataTable.Value("SAVE_NAME", dtLocalSheet)
  	
  	   REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
