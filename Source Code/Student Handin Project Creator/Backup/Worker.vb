Imports System.IO
Imports Microsoft.Win32

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public rootdirectory As String = ""
    Public department As String = ""
    Public course As String = ""
    Public project As String = ""
    Public opendate As Date
    Public closedate As Date
    Public refresh As Boolean = False

    
    Private foldercount As Long = 0

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal FolderCreated As Long)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), ) & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin Project Creator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), ) & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try
            foldercount = 0

            RaiseEvent WorkerProgress(foldercount)
            Dim dinfo As DirectoryInfo = New DirectoryInfo(rootdirectory)

            Dim dinfo2 As DirectoryInfo

            Dim yearfound As Boolean = False

            For Each dinfo2 In dinfo.GetDirectories
                If dinfo2.Name = Format(Now(), "yyyy") Then
                    yearfound = True
                    Exit For
                End If
            Next
            If yearfound = False Then
                dinfo = New DirectoryInfo(rootdirectory)
                dinfo.CreateSubdirectory(Format(Now(), "yyyy"))

            End If

            dinfo = New DirectoryInfo((rootdirectory & "\" & Format(Now(), "yyyy") & "\" & department).Replace("\\", "\"))
            If dinfo.Exists = False Then
                dinfo.Create()
            End If
            dinfo = New DirectoryInfo((rootdirectory & "\" & Format(Now(), "yyyy") & "\" & department & "\" & course).Replace("\\", "\"))
            If dinfo.Exists = False Then
                dinfo.Create()
            End If
            dinfo = New DirectoryInfo((rootdirectory & "\" & Format(Now(), "yyyy") & "\" & department & "\" & course & "\" & project).Replace("\\", "\"))
            If dinfo.Exists = False Then
                dinfo.Create()
            End If

            If refresh = False Then
                Dim filewriter As StreamWriter = New StreamWriter((rootdirectory & "\" & Format(Now(), "yyyy") & "\" & department & "\" & course & "\" & project & "\access.ini").Replace("\\", "\"), False)
                filewriter.WriteLine("open:" & (Format(opendate, "yyyy/MM/dd HH:mm")).Replace("/", "").Replace(":", "").Replace(" ", ""))
                filewriter.Write("close:" & (Format(closedate, "yyyy/MM/dd HH:mm")).Replace("/", "").Replace(":", "").Replace(" ", ""))
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            
            Dim apptorunn As String
            apptorunn = """" & (Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\") & """ """ & (rootdirectory & "\" & Format(Now(), "yyyy") & "\" & department & "\" & course & "\" & project & "\access.ini").Replace("\\", "\") & """ r "".Students_G.com.main.uct"""
            Activity_Logger(ApplicationLauncher(apptorunn))

            Dim apptorun As String = ""
            apptorun = """" & (Application.StartupPath & "\Commerce Courses Student List Extractor.exe").Replace("\\", "\") & """ """ & course & """"


            Dim students As String() = ApplicationLauncher(apptorun).Split(vbCrLf)
            For Each str As String In students
                Dim student As String = str.Trim.ToUpper.Replace(course.ToUpper & "_G ", "")
                If student.StartsWith("STATUS") = True Then
                    Exit For
                End If
                dinfo2 = New DirectoryInfo((dinfo.FullName & "\" & student).Replace("\\", "\"))
                If dinfo2.Exists = False Then
                    foldercount = foldercount + 1
                    RaiseEvent WorkerProgress(foldercount)
                    dinfo2.Create()
                    apptorun = """" & (Application.StartupPath & "\SETTRUST.exe").Replace("\\", "\") & """ """ & (dinfo.FullName & "\" & student).Replace("\\", "\") & """ rwcemf ""." & student & ".Students.com.main.uct"""
                    Activity_Logger(ApplicationLauncher(apptorun))
                End If

            Next
            dinfo2 = Nothing
            dinfo = Nothing
            result = "Success"
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub


    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            'myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If



            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Logger(ex, "DosShellCommand")
            Error_Handler(ex, "DosShellCommand")
        End Try

        Return s
    End Function

    Private Function ApplicationLauncher(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process


            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Logger(ex, "ApplicationLauncher")
            Error_Handler(ex, "ApplicationLauncher")
        End Try
        Return s
    End Function

    Public Function ReturnRegKeyValue(ByVal MainKey As String, ByVal RequestedKey As String, ByVal Value As String) As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If continue = False Then
                        Exit For
                    End If
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                result = regkey.GetValue(Value)
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            result = "Fail. Could not locate Value within Registry Key"
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "ReturnRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "ReturnRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function

    Public Function SetRegKeyValue(ByVal MainKey As String, ByVal RequestedKey As String, ByVal Value As String, ByVal Data As String) As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If continue = False Then
                        Exit For
                    End If
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                'result = regkey.GetValue(Value)
                                regkey.SetValue(Value, Data)
                                result = "Success."
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            result = "Fail. Could not locate Value within Registry Key"
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "SetRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "SetRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function


    Public Function CreateSubRegKey(ByVal MainKey As String, ByVal RequestedKey As String, Optional ByVal Value As String = "", Optional ByVal Data As String = "") As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                    If continue = False Then
                        regkey = regkey.CreateSubKey(stri)
                        continue = True
                    End If
                Next
                If Not Value = "" Then
                    result = "Success."
                End If
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                'result = regkey.GetValue(Value)
                                regkey.SetValue(Value, Data)
                                result = "Success."
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            If Not Value = "" Then
                                regkey.SetValue(Value, Data)
                            End If
                            result = "Success."
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "SetRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "SetRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function


    Public Function MapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Try
            Dim continue As Boolean = True
            Dim letterlist As ArrayList = New ArrayList
            letterlist.Clear()
            Dim i As Integer
            For i = 65 To 90
                letterlist.Add(Chr(i).ToString)
            Next
            'MsgBox(letterlist.Count)
            Dim runner As IEnumerator
            Dim fso As New Scripting.FileSystemObject
            runner = fso.Drives.GetEnumerator()
            While runner.MoveNext() = True
                Dim d As Scripting.Drive
                d = runner.Current()
                '   MsgBox(d.DriveLetter.ToString)
                letterlist.RemoveAt((letterlist.IndexOf(d.DriveLetter.ToString)))
            End While
            If letterlist.Count > 1 Then
                letterlist.RemoveAt((letterlist.IndexOf("B")))
            End If
            resultdrive = ""
            For Each strp As String In letterlist
                '  MsgBox(strp)
                resultdrive = strp
                Exit For
            Next
            If resultdrive.Trim = "" Then
                resultdrive = "Fail. No available drive letter can be found"
                continue = False
            End If

            If continue = True Then
                Dim apptorun As String
                'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
                'net use T: /DELETE
                If pathtomap.EndsWith("\") Then
                    pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
                End If
                apptorun = "net use " & resultdrive.Trim.ToUpper & ": """ & pathtomap & """ /PERSISTENT:NO"
                'MsgBox(apptorun)
                result = DosShellCommand(apptorun)
                If Not result.IndexOf("The command completed successfully.") = -1 Then
                    resultdrive = resultdrive & ":\"
                Else
                    resultdrive = "Fail. Unable to map the give path to a directory"
                End If
            End If
        Catch ex As Exception
            Error_Logger(ex, "MapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Public Function UnMapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Try
            Dim apptorun As String
            'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
            'net use T: /DELETE

            If pathtomap.EndsWith("\") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            If pathtomap.EndsWith(":") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            apptorun = "net use " & pathtomap & ":  /DELETE"

            resultdrive = DosShellCommand(apptorun)
        Catch ex As Exception
            Error_Logger(ex, "UnMapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Private Sub Error_Logger(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then

                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                If identifier_msg = "" Then
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & ex.ToString)
                Else
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                End If

                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin Project Creator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


End Class
