'############################################## Test Suite Driver #########################################################
'Author: Venkata Rajulapati
'Date:  
'Date of Modification:
'Last Modified By:
'Comments:
'#####################################################################################################################
'On Error Resume Next

'Dim TestcaseDriverData

'****** Fetching RootPath

gbl_BasePath= AutomationSuitePath
  
'******** ResourcePaths

gbl_FunctionLibraryPath=gbl_BasePath & "Lib"

gbl_ObjectRepPath=gbl_BasePath & "ObjectRepository"

gbl_FunctionalDriverpath=gbl_BasePath & "Configuration\Functional_Driver.xls"

gbl_ConfigFolderPath=gbl_BasePath & "Configuration\"

gbl_ConfigFilePath=gbl_BasePath & "Configuration\Configurations.xls"

gbl_TestDataPath=gbl_BasePath & "Database\"

'************************************************** Loading Function Libraries
'
'sqlStatement="Select * From [FunctionLibraries$] Where TobeIncluded='Yes'"
'
'Set ConfigDriverdata=GetRecordset(sqlStatement,gbl_ConfigFilePath)
'
''msgbox ObjectRepDriverdata.Recordcount
'
' ConfigDriverdata.MoveFirst
'
'If ConfigDriverdata.RecordCount>0 Then
'
'	For i=1 to ConfigDriverdata.RecordCount
'
'          str_LibName=ConfigDriverdata("Library_Name")
'
'          ExecuteFile gbl_FunctionLibraryPath&"Lib\"&str_LibName&".vbs"
'
'		  If Err.Number<>0 Then
'
'			   Reporter.ReportEvent micFail,"Associating Library Failed "&str_RepName,Err.Description
'
'			   Err.Clear
'
'		  End If
'
'          ConfigDriverdata.MoveNext
'
'	Next
'
'else
'
'    Reporter.ReportEvent micWarning,"Associating  Library","There are no Library names available check your Config Sheet"
'	
'End If
'
''**** Release Record set
'
'Set ConfigDriverdata=Nothing

'*********************************** Associating Repositories

Set qtApp=CreateObject("QuickTest.Application")

sqlStatement="Select * From [ObjectRep$] Where TobeIncluded='Yes'"

Set ConfigDriverdata=GetRecordset(sqlStatement,gbl_ConfigFilePath)

 ConfigDriverdata.MoveFirst

If ConfigDriverdata.RecordCount>0 Then

	For i=1 to ConfigDriverdata.RecordCount

          str_RepName=ConfigDriverdata("Repository_Name")

		  qtApp.Test.Actions("Action1").ObjectRepositories.Add gbl_ObjectRepPath &"\" & str_RepName &".tsr"

		  If Err.Number<>0 Then

			   Reporter.ReportEvent micFail,"Associating repositoryFailed "&str_RepName,Err.Description

			   ExitTest

			   Err.Clear

		  End If

          ConfigDriverdata.MoveNext

	Next

else

    Reporter.ReportEvent micWarning,"Associating Repositories","There are no repository names available check your Config Sheet"
	
End If

Set ConfigDriverdata=Nothing

'******** Connecting to Functional Driver

strSQL = "Select [TestCase_Master$].BS_ID as BS_ID, [TestCase_Master$].TS_ID as TS_ID, " & _
        "[TestCase_Master$].TC_ID as TC_ID, [TestCase_Master$].HostName as HostName, [TestCase_Master$].TestCase_Name as TestCase_Name, [TestCase_Master$].Test_Type as Test_Type," &_
  "[TestCase_Master$].TobeExecuted as TobeExecuted, [Business_Flow$].* " & _
    "From [TestCase_Master$] Left Outer Join " & _
        "[Business_Flow$] " & _
    "On [TestCase_Master$].BS_ID = [Business_Flow$].BS_ID And " & _
                "[TestCase_Master$].TS_ID = [Business_Flow$].TS_ID And " & _
                "[TestCase_Master$].TC_ID = [Business_Flow$].TC_ID " &_
    "Where TobeExecuted = 'Y'" 

    functionaldriverpath=gbl_ConfigFolderPath&"Functional_Driver.xls"

    Set FunctionalDriverData= GetRecordset(strSQL,functionaldriverpath)

    FunctionalDriverData.MoveFirst

    For i=1 to  FunctionalDriverData.RecordCount
 
        BS_ID=FunctionalDriverData("BS_ID")

       TS_ID=FunctionalDriverData("TS_ID")

       TC_ID=FunctionalDriverData("TC_ID")

	   TCName=FunctionalDriverData("TestCase_Name")

     '******** TestData Query  

     SqlQueryTd= "Select * From [Flight$] Where ToBeExecuted='Yes'"

     Set TestcaseDriverData= GetRecordset(SqlQueryTd,gbl_TestDataPath & TCName & ".xls")

    Call RunScenario()

    FunctionalDriverData.MoveNext

   Next


Public Function RunScenario()

  On Error Resume Next

  bl nFlag=True

   TestcaseDriverData.Filter="BS_ID='" & BS_ID & "'and TS_ID='" & TS_ID & "'and TC_ID='" & TC_ID &"'"

   TestcaseDriverData.MoveFirst

    For j=1 to TestcaseDriverData.RecordCount

          For m= 11 to FunctionalDriverData.Fields.Count

               strKeyword=FunctionalDriverData.Fields.Item(m).Value

               Reporter.ReportEvent micPass,"Running" & "_" & strkeyword,""

			   If Ucase(strKeyword)<> "END" and Not isNull(strKeyword) Then

                      If Eval ("UDF_"& strKeyword & "(TestcaseDriverData )") <>True Then

                            blnFlag=False
				            
				            Reporter.ReportEvent micFail,strkeyword, " Failed"

							Reporter.ReportEvent micFail,BS_ID & "_" & TS_ID & "_" & TC_ID & "_" & TCName,"Failed"

							Exit For

					 Else

					        Reporter.ReportEvent micPass, strkeyword, " Passed"  

			        End if

		     End if

              If Ucase(strKeyword) ="END" Then

                	   Reporter.ReportEvent micPass,BS_ID & "_" & TS_ID & "_" & TC_ID & "_" & TCName,"Passed"

					   Exit For

			  End If

			  If isnull(strKeyword) Then

                  Reporter.ReportEvent micFail,BS_ID & "_" & TS_ID & "_" & TC_ID & "_" & TCName,"Failed since keyword is empty check business flow sheet"

				  Exit For

			  End If

	   Next

	   TestcaseDriverData.MoveNext

     Next

    Set FlightTestData=Nothing

End Function





							
