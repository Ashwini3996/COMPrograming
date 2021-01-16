Imports System.Windows.Forms

'.Net calleble dll created by using tlbimp.exe utility
Imports AutomationServerTypeLibForDotNet

Public Class VBAutomation
	Inherits Form
	Public Sub New()
		Dim MyDispatch As Object
		Dim MyRef As New CMyMathClass
		MyDispatch = MyRef
        Dim iNum1 = 175
        Dim iNum2 = 125
		Dim iSum = MyDispatch.SumOfTwoIntegers(iNum1, iNum2)
		Dim str As String = String.Format("Sum of {0} And {1} Is {2}", iNum1, iNum2, iSum)
        'default message box with only 1 button of 'Ok'
        MsgBox(str)
		
		Dim iSub =  MyDispatch.SubtractionOfTwoIntegers(iNum1, iNum2)
		 str = String.Format("Subtraction of {0} And {1} Is {2}", iNum1, iNum2, iSub)   
        MsgBox(str)
		
		'Following statment i.e 'End' Works as DistroyWindow(hwnd)
		End
	End Sub 

    <STAThread()>
	Shared Sub Main()
		Application.EnableVisualStyles()
		Application.Run(New VBAutomation())
	    End Sub
	End Class  	
		
	
	