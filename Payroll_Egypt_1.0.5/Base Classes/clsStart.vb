Public Class clsStart
    
    Shared Sub Main()
        Dim oRead As System.IO.StreamReader
        Dim LineIn, strUsr, strPwd As String
        Dim i As Integer
        Try
            Try
                oApplication = New clsListener
                oApplication.Utilities.Connect()
                oApplication.SetFilter()
                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                    LocalCurrency = .GetAdminInfo.LocalCurrency
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try

            If 1 = 2 Then ' oApplication.Utilities.CheckLicense() = False Then
                System.Windows.Forms.Application.Exit()
            Else

                oApplication.Utilities.CreateTables()
                oApplication.Utilities.AddRemoveMenus("Menu.xml")
                Dim omenuItem As SAPbouiCOM.MenuItem
                omenuItem = oApplication.SBO_Application.Menus.Item("Z_mnu_Pay601")
                omenuItem.Image = Application.StartupPath & "\Inv.bmp"

                oApplication.Utilities.createPayrollMainAuthorization()
                oApplication.Utilities.AuthorizationCreation()
                oApplication.Utilities.Message("Payroll Egypt Addon Connected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oApplication.Utilities.NotifyAlert()
                System.Windows.Forms.Application.Run()
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
