'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New PhishingRibbon()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports Microsoft.Office.Interop.Outlook

<Runtime.InteropServices.ComVisible(True)> _
Public Class PhishingRibbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub ForwardActiveItem()
        Dim objMail As Outlook.MailItem
        Dim objItem = GetCurrentItem()
        Dim PropName, Header As String
        Dim oPA As Outlook.PropertyAccessor
        If MsgBox("Thank you for reporting this suspicious email. If you clicked on any link or attachment within this email, please notify us at " &
               "PhishingNotice@GAIG.com. If you did not click any link or attachment, you may consider this situation closed. The suspicious " &
               "email has been moved to your deleted items folder.", vbOKCancel + vbInformation, "Phish Reporter") = vbCancel Then
            Exit Sub
        Else
            objMail = objItem.Forward
            PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
            oPA = objItem.PropertyAccessor
            Header = oPA.GetProperty(PropName)
            objMail.Body = Header & vbCrLf & vbCrLf & "Email Body: " & vbCrLf & objItem.Body
            objMail.Subject = "Suspected Phishing Attempt " & objItem.Subject
            objMail.Recipients.Add("tthrockmorton@gaig.com") 'You add the name associated with the email or the email address here - send to PhishingNotice@GAIG.com
            objMail.Save()
            objMail.Send() 'Forwards email to specified recipient
            objItem.Delete() 'Deletes original email
        End If
        objMail = Nothing
        objItem = Nothing 'Nulls out variables to clear out memory
    End Sub

    'Public Sub dbConnect()
    '    Dim password As String
    '    Dim sqlStr As String
    '    Dim server_Name As String
    '    Dim user_ID As String
    '    Dim database_Name As String
    '    Dim cn As Object
    '    Dim rs As Object
    '    rs = CreateObject("ADODB.Recordset")
    '    server_Name = "http://localhost:3306"
    '    user_ID = "phishReport"
    '    database_Name = "gaig_users"
    '    password = "#KjW9#Q8Fpt1PC8YnMrd1e"
    '    cn = CreateObject("ADODB.Connection")
    '    cn.Open("Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & server_Name & ";Database=" & database_Name & ";Uid=" & user_ID & ";Pwd=" & password & ";")

    'End Sub

    'Uses cases to select the currently inspected item when viewing a folder or the item viewed if you have a specific item open.
    Function GetCurrentItem() As Object
        Dim objApp As Outlook.Application
        objApp = New Application
        On Error Resume Next
        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer" 'If looking at a folder structure and previewing one item
                GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
            Case "Inspector" 'You have a specific item opened
                GetCurrentItem = objApp.ActiveInspector.CurrentItem
            Case Else
        End Select
        objApp = Nothing
    End Function

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("PhishingAddinOutlook2010.PhishingRibbon.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub OnAction(ByVal control As Office.IRibbonControl)
        If (control.Id = "textButton") Then
            ForwardActiveItem()
        End If
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
