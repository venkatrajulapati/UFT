'################################################################## Function Libary #####################################################
'File Name                                                  Common Function
'Author                                                       Venkata Rajulapati
'Createddate                                               03-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'Naming Convensions                           UDF_SCF_FunctionName
'####################################################################################################################################


'############################################################################### #####################################################
'FunctionName                                          UDF_SCF_GetRecordset
'Author                                                       Venkata Rajulapati
'Createddate                                               03-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'####################################################################################################################################
Public Function GetRecordset(sqlStatement,strFilepath)
   '*******Create DB Object
   Set oDbConnection=CreateObject("ADODB.Connection")
   oDbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strFilepath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=0\"";"  
   'oDbConnection.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & strFilepath & ";Readonly=True"
   If Err.Number<>0 Then
   
	   Reporter.ReportEvent micFail,"Connecting to Excel as DB","Connection Failed"
	   Err.Clear
	   ExitTest
	else
   '******* Create Record set
   Set oRecordset=CreateObject("ADODB.Recordset")
 '  oRecordset.CursorLocation=3
   oRecordset.Open sqlStatement,oDbConnection,1,3
  
   If Err.Number<>0 Then
      Reporter.ReportEvent micFail,"Creating Recordset","Recordset Failed"
	  Err.clear
     ExitTest
   else
	 Set GetRecordset=oRecordset.Clone
	 Set oRecordset=Nothing
	 Set oDbConnection=Nothing
   End If
End if
End Function
