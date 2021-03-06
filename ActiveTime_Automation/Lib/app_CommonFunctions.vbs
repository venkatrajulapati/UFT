'################################################################## Function Libary #####################################################
'File Name                                                  AppCommon Function
'Author                                                       Venkata Rajulapati
'Createddate                                               26-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'Naming Convensions                           UDF_SCF_FunctionName
'####################################################################################################################################

'############################################################################### #####################################################
'FunctionName                                          UDF_Login
'Author                                                       Venkata Rajulapati
'Createddate                                               26-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'####################################################################################################################################
'ExecuteFile "C:\Stabilization\Test_Suite\Application\MercuryTours\Flight\Lib\globalvariables.vbs"
Public Function UDF_LaunchApplication(TestcaseDriverData)
 On Error Resume Next
'***** Variables
blnFlag=True

SystemUtil.CloseDescendentProcesses
SystemUtil.Run "C:\Program Files (x86)\HP\QuickTest Professional\samples\flight\app\flight4a.exe"

'If Window("Win_FlightReservation").Exist(3) Then
'blnFlag=True
'else'
'blnFlag=False
'End If

If Err.Number<>0 Then
   blnFlag=False
   Err.Clear
End If

If Not blnFlag Then
	UDF_LaunchApplication=False
else
   UDF_LaunchApplication=True
End If

End Function
'############################################################################### #####################################################
'FunctionName                                          UDF_Login
'Author                                                       Venkata Rajulapati
'Createddate                                               26-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'####################################################################################################################################
'ExecuteFile "C:\Stabilization\Test_Suite\Application\MercuryTours\Flight\Lib\globalvariables.vbs"
Public Function UDF_Login(TestcaseDriverData)
 On Error Resume Next
'***** Variables
blnFlag=True
wait(2)
UserName=TestcaseDriverData("UserName").Value
Password=TestcaseDriverData("Password").Value

'*******Enter UserName and Password
If Dialog("Dlg_Login").Exist(2) Then

   Dialog("Dlg_Login").Activate
   Dialog("Dlg_Login").WinEdit("Txt_AgentName").Click
   Dialog("Dlg_Login").WinEdit("Txt_AgentName").Set UserName
   Dialog("Dlg_Login").WinEdit("Txt_Password").Set Password
   Dialog("Dlg_Login").WinButton("Btn_OK").Click

End if
wait(4)
'If Window("Win_FlightReservation").Exist(3) Then
'blnFlag=True
'else'
'blnFlag=False
'End If

If Err.Number<>0 Then
   blnFlag=False
   Err.Clear
End If

If Not blnFlag Then
	UDF_Login=False
else
   UDF_Login=True
End If

End Function
'############################################################################### #####################################################
'FunctionName                                          UDF_Enter_ReservationDetails
'Author                                                       Venkata Rajulapati
'Createddate                                               26-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'####################################################################################################################################
Public Function UDF_Enter_ReservationDetails(TestcaseDriverData)
    On Error Resume Next
blnFlag=True
'***** Variables
Date_of_Flight=TestcaseDriverData("Date_of_Flight")
Fly_From=TestcaseDriverData("Fly_From")
Fly_To=TestcaseDriverData("Fly_To")
Passenger_Name=TestcaseDriverData("Passenger_Name")
Reservation_Class=TestcaseDriverData("Reservation_Class")
Noof_Tickets=TestcaseDriverData("Noof_Tickets")

'*******Enter UserName and Password
If Window("Win_FlightReservation").Exist(2) Then
'******Setting Date Field
  If Not isNull(Date_of_Flight) Then
      	   Window("Win_FlightReservation").WinObject("Date of Flight:").Type Date_of_Flight
  End If
'***** select FlyFrom   
 If Not isNull(Fly_From) Then
      	   If Window("Win_FlightReservation").WinComboBox("Cmb_Fly From").Exist Then
			   window("Win_FlightReservation").WinComboBox("Cmb_Fly From").Select Fly_From
		   End If
  End If
End if
'***** select FlyTo
 If Not isNull(Fly_To) Then
      	   If Window("Win_FlightReservation").WinComboBox("Cmb_FlyTo").Exist Then
			   window("Win_FlightReservation").WinComboBox("Cmb_FlyTo").Select Fly_To
		   End If
  End If

'***** Click on Flight Button

If Window("Win_FlightReservation").WinButton("Btn_Flight").Exist Then
	Window("Win_FlightReservation").WinButton("Btn_Flight").Click
	Window("Win_FlightReservation").Dialog("Flights Table").WinButton("OK").Click
End If

'***** Enter the Name of Passanger

If Not isNull(Passenger_Name) Then

	If Window("Win_FlightReservation").WinEdit("Txt_Name").Exist Then
		  Window("Win_FlightReservation").WinEdit("Txt_Name").Set Passenger_Name
	End If
End If

'***** Select Reservation Class

If Not isNull(Reservation_Class) Then

Select Case Reservation_Class

Case "First"
	Window("Win_FlightReservation").WinRadioButton("First").Set "ON"

Case "Economy"
	Window("Win_FlightReservation").WinRadioButton("Economy").Set "ON"

Case "Business"
	Window("Win_FlightReservation").WinRadioButton("Business").Set "ON"
End Select
	
End If


If Err.Number<>0 Then

  g_ErrMsg=Err.Description
  Err.Clear
   blnFlag=False

End If

If Not blnFlag Then
	UDF_Enter_ReservationDetails=False
	'Reporter.ReportEvent micFail,"Enter reservation Deatils","Enter Details failed " & g_ErrMsg
else
   UDF_Enter_ReservationDetails=True
   'Reporter.ReportEvent micPass,"Enter reservation Deatils","Enter Details passed " 
End If

End Function
'############################################################################### #####################################################
'FunctionName                                          UDF_Nav_Insertorder
'Author                                                       Venkata Rajulapati
'Createddate                                               26-May-2014
'Modification History:                      
'Modified By:
'Modified Date
'####################################################################################################################################
Public Function UDF_Nav_Insertorder()
   blnFlag=True

   If Window("Win_FlightReservation").WinButton("Btn_InsertOrder").Exist Then
	   Window("Win_FlightReservation").WinButton("Btn_InsertOrder").Click
   End If

If Err.Number<>0 Then
	 g_ErrMsg=Err.Description
	 Err.Clear
	 blnFlag=False
End If

If Not blnFlag Then
	Reporter.ReportEvent micFail,"Click on Insert Order","Not Clicked "& g_ErrMsg
	Reporter.ReportEvent micPass,"Click on Insert Order","Clicked"

End If
End Function




'Public Function RunScenario()
'blnFlag=True
'   TestcaseDriverData.Filter="BS_ID='" & BS_ID & "'and TS_ID='" & TS_ID & "'and TC_ID='" & TC_ID &"'"
'   TestcaseDriverData.MoveFirst
'      For j=1 to TestcaseDriverData.RecordCount
'          For m= 13 to FunctionalDriverData.Fields.Count
'               strKeyword=FunctionalDriverData.Fields.Item(m).Value
'             ' Call UDF_RunScenario(strKeyword,TestcaseDriverData)
'			  If Eval ("UDF_"& strKeyword & "( )") <>True Then
'                  blnFlag=trFalse
'				  Exit For
'				  Reporter.ReportEvent micFail,"RunningKeywrod", strKeyword & " Failed"
'			  End if
'			    
'
'			   If Ucase(strKeyword)="END" Then
'				   Exit For
'			   End If
'	      Next
'	   TestcaseDriverData.MoveNext
'     Next
'    Set FlightTestData=Nothing
'End Function
