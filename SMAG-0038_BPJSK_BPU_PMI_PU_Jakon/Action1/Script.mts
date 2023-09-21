REM - BPJSK Automation test
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim dt_Type, dt_PayCode, dt_SaveName, dt_PinTrx
Dim dt_Username, dt_Password

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("KeagenanMobile_Lib_Report.xlsx", "SMAG-0038_BPJSK_BPU_PMI_PU_Jakon.xlsx", "SMAG-0038")
Call spGetDatatable()

Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Automation Testing : " & dt_TestScenarioDesc))


REM ------- Keagenan Mobile
Call Login(dt_Username, dt_Password)
Call Open_MAgen() 
Call Goto_BpjsK()

If dt_TCID = "MAG0038-003" or dt_TCID = "MAG0038-019" or dt_TCID = "MAG0038-035" or dt_TCID = "MAG0038-051" Then 
	 Call Choice_Fav(dt_TCID)
	 Call Choice_Periode()
'	Msgbox("Pilih Fav")
Else
	Call Input_Baru(dt_TCID, dt_Type, dt_PayCode, dt_SaveName)
End If

If dt_TCID = "MAG0038-005" or dt_TCID = "MAG0038-007" or dt_TCID = "MAG0038-008" or dt_TCID = "MAG0038-021" or dt_TCID = "MAG0038-023" or dt_TCID = "MAG0038-024" or dt_TCID = "MAG0038-037" or dt_TCID = "MAG0038-039" or dt_TCID = "MAG0038-040" Then
	Call Verify_Fail()
Else
	Call Input_Pass(dt_TCID, dt_PinTrx)
End If

If dt_TCID = "MAG0038-004" or dt_TCID = "MAG0038-006" or dt_TCID = "MAG0038-020" or dt_TCID = "MAG0038-022" or dt_TCID = "MAG0038-036" or dt_TCID = "MAG0038-038" or dt_TCID = "MAG0038-052" or dt_TCID = "MAG0038-054" Then
	Call Verify_Fail()
ElseIf dt_TCID = "MAG0038-002" or dt_TCID = "MAG0038-018" or dt_TCID = "MAG0038-034" or dt_TCID = "MAG0038-050" Then
	Call Verify_Success_Save()
Else
	Call Verify_Success()
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
	
	REM ---- BPJSK Keagenan Mobile Library
	LoadFunctionLibrary (LibPathKeagenanMobile & "Lib_KeagenanMobile_BPJSK.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_KeagenanMobile_BPJSK.tsr")
End Sub

Sub spGetDatatable()
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

	REM --------- Pulsa Prabayar
	dt_Type		        = DataTable.Value("TYPE", dtLocalSheet)
	dt_PayCode		    = DataTable.Value("NIK_KODEBAYAR", dtLocalSheet)
	dt_PinTrx	    	= DataTable.Value("PIN_TRX", dtLocalSheet)
    dt_SaveName	    	= DataTable.Value("SAVE_NAME", dtLocalSheet)
    
    	   	REM --------- Login
	dt_Username		= DataTable.Value("USERNAME", dtLocalSheet)
	dt_Password		= DataTable.Value("PASSWORD", dtLocalSheet)
End Sub
