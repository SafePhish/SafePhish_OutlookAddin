Imports Microsoft.Office.Core

Public Class PhishingAddin

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As IRibbonExtensibility
        Return New PhishingRibbon()
    End Function

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
