
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports VB6 = Microsoft.VisualBasic
Imports System.Security.Principal
Imports System.IO.Compression
Imports System.IO
Imports IWshRuntimeLibrary
Imports System.Net
Imports System.ComponentModel
Imports System.Xml

Public Class HtaInstall
    Const JSBURL = "http://jsonbasic.azurewebsites.net/hta/"
    Const ZIPURL = "http://jsonbasic.azurewebsites.net/___JSBHTML5/"
    Const ShortCutName = "JSB (hta)"
    Dim InStallDir As String = "C:\jsb\"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' elevate ourselves
        If Not Debugger.IsAttached Then
            If ElevateOurSelves() = False Then End
        End If

        Me.ProgressBar1.Maximum = 60

        Me.Show()
        Me.ProgressBar1.Value = 10
        Application.DoEvents()

        ' Get and process ZIP file
        If Not fetchZipFile() Then End

        Me.ProgressBar1.Value = 20

        If Not unZipFiles() Then End

        Me.ProgressBar1.Value = 30

        ' Get and process DLL file
        If Not fetchDLLFiles() Then End

        Me.ProgressBar1.Value = 40

        RegisterDllFiles()

        Me.ProgressBar1.Value = 50

        ' Create a Deskop ICON
        Dim DeskTopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim StartMenuFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)

        CreateShortCut(ShortCutName, InStallDir & "tcl.hta", DeskTopFolder, "", "%SystemRoot%\system32\SHELL32.dll", 18, "")
        CreateShortCut(ShortCutName, InStallDir & "tcl.hta", StartMenuFolder, "", "%SystemRoot%\system32\SHELL32.dll", 18, "")

        Me.ProgressBar1.Value = 60

        Dim hta As Object = Nothing

        Try
            hta = CreateObject("dbADO.clsHtaConnection")
        Catch ex As Exception
            MsgBox("I was unable to create the necessary dbADO.dll object (dbADO.clsHtaConnection)" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End
        End Try

        ' Attempt to create object and test
        Try
            hta.adoInit(InStallDir & "tcl.hta", "hta")
        Catch ex As Exception
            MsgBox("I was unable to call adoInit() in dbADO.clsHtaConnection" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End
        End Try

        ' Start the application?
        Dim proc2 As New System.Diagnostics.Process()
        proc2 = Process.Start(InStallDir & "tcl.hta", "")

        Application.Exit()  ' Quit itself
    End Sub

    Function ElevateOurSelves() As Boolean
        If AreWeElevated() Then Return True

        Dim proc As New ProcessStartInfo
        proc.UseShellExecute = True
        proc.WorkingDirectory = Environment.CurrentDirectory
        proc.FileName = Application.ExecutablePath
        proc.Verb = "runas"

        Try
            Process.Start(proc)
        Catch
            ' The user refused to allow privileges elevation.
            ' Do nothing and return directly ...
        End Try

        Application.Exit()  ' Quit itself
        Return False
    End Function

    Dim doneLoadDone As Boolean = False

    Function fetchZipFile()
        Dim destPath As String = Path.GetTempPath() & "__HTA.zip"

        Me.Text = "Downloading ZIP File"
        Try
            If System.IO.File.Exists(destPath) Then My.Computer.FileSystem.DeleteFile(destPath)
        Catch ex As Exception
            MsgBox("fetchZipFile fail: " & Err.Description)
            Return False
        End Try

        Try
            Dim myWebClient As New System.Net.WebClient()
            AddHandler myWebClient.DownloadProgressChanged, AddressOf DownloadProgressCallback
            AddHandler myWebClient.DownloadFileCompleted, AddressOf DownloadFileCallback2

            doneLoadDone = False
            myWebClient.DownloadFileAsync(New Uri(ZIPURL & "__HTA.zip"), destPath)
        Catch ex As Exception
            MsgBox("fetchZipFile __HTA.zip failed: " & Err.Description, MsgBoxStyle.Critical)
            Return False
        End Try

        Do Until doneLoadDone
            Application.DoEvents()
        Loop
        Return True
    End Function

    Sub DownloadProgressCallback(sender As Object, args As DownloadProgressChangedEventArgs)
        Me.Text = Split(Me.Text, "File")(0) & "File " & args.ProgressPercentage & "%"
    End Sub

    Sub DownloadFileCallback2(sender As Object, args As AsyncCompletedEventArgs)
        doneLoadDone = True
    End Sub

    Function fetchDLLFiles()

        Me.Text = "Downloading DLL File"
        If wipeOut("dbADO") = False Then Return False

        Try
            Dim myWebClient As New System.Net.WebClient()
            AddHandler myWebClient.DownloadProgressChanged, AddressOf DownloadProgressCallback
            AddHandler myWebClient.DownloadFileCompleted, AddressOf DownloadFileCallback2




            myWebClient.DownloadFileAsync(New Uri(JSBURL & "dbADO.dll"), InStallDir & "dbADO.dll")
        Catch ex As Exception
            MsgBox("fetchDLLFile fail: " & Err.Description, MsgBoxStyle.Critical)
            Return False
        End Try

        Do Until doneLoadDone
            Application.DoEvents()
        Loop
        Return True
    End Function

    Function wipeOut(dllName)
        If My.Computer.FileSystem.FileExists(InStallDir & dllName & ".dll") Then
            Dim I As Integer = 1
            Do While My.Computer.FileSystem.FileExists(InStallDir & "x" & dllName & I & ".dll")
                I = I + 1
            Loop

            Try
                My.Computer.FileSystem.RenameFile(InStallDir & dllName & ".dll", "x" & dllName & I & ".dll")
            Catch ex As Exception
                MsgBox("fetchDLLFile rename fail: " & Err.Description, MsgBoxStyle.Critical)
                Return False
            End Try
        End If
        Return True
    End Function

    Function RegisterDllFiles() As Boolean
        Me.Text = "Registering DLL Files"

        RegisterDllFile("dbADO.dll")

        RegisterDllFile("System.Data.SQLite.dll") ' Doesn't matter if we error if it's already installed
        RegisterDllFile("System.Data.SQLite.Linq.dll") ' Doesn't matter if we error if it's already installed
        RegisterDllFile("System.Data.SQLite.EF6.dll") ' Doesn't matter if we error if it's already installed

        Return True
    End Function

    Function RegisterDllFile(ByVal dllName As String) As Boolean
        Dim Asm As Assembly = Nothing

        Try
            Asm = Assembly.LoadFile(InStallDir & dllName)
        Catch ex As Exception
            MsgBox("LoadFile " & dllName & " failed: " & Err.Description)
            Return False
        End Try

        Dim regAsm As RegistrationServices = Nothing
        Try
            regAsm = New RegistrationServices()
        Catch ex As Exception
            MsgBox("RegistrationServices fail: " & Err.Description)
            Return False
        End Try

        Dim didIt As Boolean = False
        Try
            didIt = regAsm.RegisterAssembly(Asm, AssemblyRegistrationFlags.SetCodeBase)
        Catch ex As Exception
            MsgBox("RegisterDllFile fail: " & Err.Description, MsgBoxStyle.Critical)
            Return False
        End Try

        Return didIt
    End Function

    Function unZipFiles()
        Dim zipPath As String = Path.GetTempPath() & "__HTA.zip"

        Me.Text = "UnZipping ZIP File"

        Try
            Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
                For Each entry As ZipArchiveEntry In archive.Entries
                    Debug.Print(entry.FullName)

                    Dim Path As String = VB6.Replace(entry.FullName, "/", "\")
                    If VB6.Left(Path, 1) = "\" Or Mid(Path, 2, 1) = ":" Then

                    ElseIf UCase(VB6.Left(Path, 6)) = "IFWIS\" Then
                        Path = VB6.Left(InStallDir, 3) & Path

                    ElseIf UCase(VB6.Left(Path, 4)) = "JSB\" Then
                        Path = VB6.Left(InStallDir, 3) & Path

                    Else
                        Path = InStallDir & Path
                    End If

                    Dim ext = LCase(GetExtension(Path))
                    Dim folder = LCase(GetFolder(Path))

                    '  Path = Mid(Path, InStr(Path, "\") + 1)

                    If VB6.Right(Path, 1) = "\" Then
                        If (Not System.IO.Directory.Exists(Path)) Then System.IO.Directory.CreateDirectory(Path)

                        ' don't overwrite existing databases
                    ElseIf VB6.Right(folder, 10) = "\app_data\" AndAlso (ext = "accdb" Or ext = "mdb" Or ext = "mdf" Or ext = "ldf" Or ext = "db") Then
                        If System.IO.File.Exists(Path) = False Then entry.ExtractToFile(Path, False)

                    Else
                        entry.ExtractToFile(Path, True)
                    End If
                Next
            End Using
        Catch ex As Exception
            MsgBox("unZipFiles fail: " & Err.Description, MsgBoxStyle.Critical)
            Return False
        End Try

        Return True
    End Function

    Function AreWeElevated() As Boolean
        Dim identity = WindowsIdentity.GetCurrent()
        Dim principal = New WindowsPrincipal(identity)
        Return principal.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    Function AreWeRegistered() As Boolean
        Dim adoDB As Object = Nothing

        Try
            adoDB = CreateObject("dbADO.clsHTAConnection")
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Function CreateShortCut(ByVal shortcutName As String, ByVal targetFullpath As String, Optional ByVal creationDir As String = "", Optional ByVal workingDir As String = "", Optional ByVal iconFile As String = "", Optional iconNumber As Integer = 0, Optional ByVal Arguments As String = "") As Boolean

        Me.Text = "Creating Short Cut"

        Try
            If creationDir = "" Then creationDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If Not IO.Directory.Exists(creationDir) Then
                Dim retVal As DialogResult = MsgBox(creationDir & " does not exist. Do you wish to create it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                If retVal = DialogResult.Yes Then
                    IO.Directory.CreateDirectory(creationDir)
                Else
                    Return False
                End If
            End If

            Dim wShell As New WshShell
            Dim shortCut As IWshRuntimeLibrary.IWshShortcut
            shortCut = CType(wShell.CreateShortcut(creationDir & "\" & shortcutName & ".lnk"), IWshRuntimeLibrary.IWshShortcut)
            shortCut.TargetPath = targetFullpath
            shortCut.WindowStyle = 1
            shortCut.Description = shortcutName
            If workingDir <> "" Then shortCut.WorkingDirectory = workingDir Else shortCut.WorkingDirectory = GetFolder(targetFullpath)
            If iconFile <> "" Then shortCut.IconLocation = iconFile & "," & iconNumber
            If Arguments <> "" Then shortCut.Arguments = Arguments
            shortCut.Save()
            Return True
        Catch ex As System.Exception
            MsgBox("CreateShortCut fail: " & Err.Description, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

    Public Function GetFolder(ByVal SFile As String) As String
        For lCount As Integer = Len(SFile) To 1 Step -1
            If VB6.Mid$(SFile, lCount, 1) = "\" Or VB6.Mid$(SFile, lCount, 1) = "/" Then Return VB6.Left$(SFile, lCount)
        Next
        Return vbNullString
    End Function

    Public Function GetExtension(ByVal SFile As String) As String
        Dim lCount As Integer
        For lCount = Len(SFile) To 2 Step -1
            If VB6.InStr(":\/", VB6.Mid$(SFile, lCount, 1)) > 0 Then Exit For
            If VB6.Mid$(SFile, lCount, 1) = "." Then Return VB6.Mid$(SFile, lCount + 1)
        Next
        Return ""
    End Function
End Class
