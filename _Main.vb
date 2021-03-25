Public Module Main

    Private _ErrorLogLocation As String = ""

    Public Sub Main(args As String())
        AddHandler Application.ThreadException, AddressOf AprBase.Application_ThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AprBase.CurrentDomain_UnhandledException
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
        Application.EnableVisualStyles()

        Dim MainForm As Form = New frmMain()

        AprBase.Initialise(MainForm)
        AprBase.Settings.Icon = My.Resources.INFO

        AprBase.Settings.DebugLogging = AprBase.Settings.Get(Of Boolean)("DebugEnabled")
        AprBase.Settings.DebugLogging_DB = AprBase.Settings.DebugLogging
        AprBase.Settings.DebugLogging_Screen = AprBase.Settings.DebugLogging

        AprBase.Errors.Handlers.AddRange({
                                         New AprBase.AddErrorLogEventHandler(AddressOf LogError_File),
                                         New AprBase.AddErrorLogEventHandler(AddressOf LogError_Screen)
                                     })

        Application.Run(MainForm)
    End Sub

    Public Function GetErrorLogLocation() As String
        If _ErrorLogLocation.IsNotSet() Then
            _ErrorLogLocation = "{0}Applications\Errors\BDOManager.txt".FormatWith(AprBase.GetBaseDirectory(AprBase.BaseDirectory.GoogleDrive))
        End If

        Return _ErrorLogLocation
    End Function

    Private Sub LogError_File(silent As Boolean, [error] As AprBase.Models.Error)
        Try
            AprBase.IORoutines.WriteToFile(True, True, GetErrorLogLocation(), False, [error].ErrorDetails)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LogError_Screen(silent As Boolean, [error] As AprBase.Models.Error)
        If silent Then
            Return
        End If

        Try
            AprBase.MessageBox.Show([error].Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
        End Try
    End Sub

End Module