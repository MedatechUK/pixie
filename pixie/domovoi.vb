Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Xml.Schema
Imports System.Xml

Module domovoi

    '==================================
    ' Pixie://
    ' By Si.
    '
    ' Protocol for http binary launch
    ' on the localhost
    '==================================

#Region "declarations"

    Public doc As New XmlDocument

    Public rxProg As Regex = New Regex( _
        "(?<=\:\/\/)[A-Za-z0-9]*(?=\/|\?)", _
        RegexOptions.IgnoreCase _
    )

    Public rxParam As Regex = New Regex( _
        "(?<=\?|\&)[A-Za-z0-9]*(?=\=)", _
        RegexOptions.IgnoreCase _
    )

#End Region

#Region "Properties"

    Private ReadOnly Property AppPath() As String
        Get
            Static ret As String = Nothing
            If IsNothing(ret) Then
                Dim ap As New DirectoryInfo( _
                    New Uri( _
                        System.IO.Path.GetDirectoryName( _
                            System.Reflection.Assembly.GetExecutingAssembly().CodeBase _
                        ) _
                    ).LocalPath _
                )
                ret = ap.FullName

                UnPack(My.Resources.valid, ret, "valid.xsd")
                UnPack(My.Resources.help.Replace("$PATH$", ret), ret, "help.html")
                UnPack(My.Resources.example, ret, "example.txt")

            End If

            Return ret

        End Get
    End Property

    Private ReadOnly Property UserFolder() As String
        Get
            Static ret As String = Nothing
            If IsNothing(ret) Then

                Dim appFolder As New DirectoryInfo( _
                    IO.Path.Combine( _
                        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), _
                        "medatech" _
                    ) _
                )
                With appFolder
                    If Not .Exists Then .Create()
                End With

                Dim pixieFolder As New DirectoryInfo( _
                    IO.Path.Combine( _
                        appFolder.FullName, _
                        "pixie" _
                    ) _
                )
                With pixieFolder
                    If Not .Exists Then .Create()
                End With

                ret = pixieFolder.FullName
                UnPack(My.Resources.programs, ret, "programs.xml")

            End If

            Return ret

        End Get
    End Property

    Private ReadOnly Property Settings() As XmlReaderSettings
        Get
            Static ret As XmlReaderSettings = Nothing
            If IsNothing(ret) Then

                ret = New XmlReaderSettings()
                With ret
                    .Schemas.Add( _
                        "", _
                        Path.Combine( _
                            AppPath, _
                            "valid.xsd" _
                        ) _
                    )
                    .ValidationType = ValidationType.Schema
                End With
            End If

            Return ret

        End Get
    End Property

#End Region

#Region "Private Methods"

    Private Sub UnPack(ByRef Res As String, ByVal folder As String, ByVal Name As String)
        If Not File.Exists( _
            Path.Combine( _
                folder, _
                Name _
            ) _
        ) Then
            File.WriteAllText( _
                Path.Combine( _
                    folder, _
                    Name _
                ), _
                Res _
            )
        End If
    End Sub

    Private Function rxArg(ByVal param As String)
        Return New Regex( _
            String.Format( _
                "(?<={0}\=).*?(?=\&|$)", _
                param _
            ), _
            RegexOptions.IgnoreCase _
        )
    End Function

    Private Function rxMatch(ByVal Pattern As Regex, ByRef SearchString As String) As List(Of String)
        Dim ret As New List(Of String)
        Dim M As Match = Pattern.Match(SearchString)
        Do While M.Success
            If Not ret.Contains(M.Value) Then
                ret.Add(M.Value)
            End If
            M = M.NextMatch
        Loop
        Return ret
    End Function

    Private Sub ValidProg(ByVal ProgramName As String, ByRef ProgNode As XmlNode)

        With ProgNode
            If IsNothing(ProgNode) Then _
                Throw New Exception(String.Format("No configuration for program '{0}'.", ProgramName))

            If Not IsNothing(.Attributes("workingdirectory")) Then
                If Not File.Exists( _
                    Path.Combine( _
                        .Attributes("workingdirectory").Value, _
                        .Attributes("filename").Value _
                    ) _
                ) Then _
                    Throw New Exception( _
                        String.Format( _
                            "{0} executable not found [{1}].", _
                            ProgramName, _
                            Path.Combine( _
                                .Attributes("workingdirectory").Value, _
                                .Attributes("filename").Value _
                            ) _
                        ) _
                    )
            End If

        End With

    End Sub

    Private Function Arguments(ByRef ProgNode As XmlNode) As String
        Dim args As String = ProgNode.Attributes("arguments").Value
        For Each p As String In rxMatch(rxParam, EnvironmentCommandLine)
            Dim a As List(Of String) = rxMatch(rxArg(p), EnvironmentCommandLine)
            If a.Count > 0 Then
                args = Replace( _
                    args, _
                    String.Format( _
                        "%{0}", _
                        p _
                    ), _
                    a(0) _
                )
            End If
        Next
        Return args
    End Function

    Private Sub Log(ByRef sw As IO.StreamWriter, ByVal Message As String)
        Try
            sw.WriteLine( _
                String.Format( _
                    "{0}{1}{2}{1}{3}", _
                    Now.ToString, _
                    Chr(9), _
                    Environment.MachineName, _
                    Message _
                ) _
            )
        Catch : End Try
    End Sub

#End Region

#Region "Start Process"

    Private Sub StartProcess(ByVal filename As String, Optional ByVal Arguments As String = Nothing, Optional ByVal WorkingDirectory As String = Nothing)
        Dim myProcess As Process = New Process()
        With myProcess
            With .StartInfo
                .FileName = filename
                If Not IsNothing(Arguments) Then _
                    .Arguments = Arguments
                If Not IsNothing(WorkingDirectory) Then _
                    .WorkingDirectory = WorkingDirectory
                .UseShellExecute = True
                .CreateNoWindow = False
            End With
            .Start()
        End With
    End Sub

    Private Sub StartProcess(ByVal Node As XmlNode)
        With Node
            If IsNothing(.Attributes("workingdirectory")) Then
                StartProcess( _
                    .Attributes("filename").Value, _
                    Arguments(Node) _
                )
            Else
                StartProcess( _
                    .Attributes("filename").Value, _
                    Arguments(Node), _
                    .Attributes("workingdirectory").Value _
                )
            End If

        End With
    End Sub

#End Region

#Region "Main"

    Private EnvironmentCommandLine As String

    Sub Main()

        EnvironmentCommandLine = "pixie:" & Split(Environment.CommandLine, "pixie:", , CompareMethod.Text)(1)
        EnvironmentCommandLine = EnvironmentCommandLine.Substring(0, EnvironmentCommandLine.Length - 1)
        EnvironmentCommandLine = EnvironmentCommandLine.Replace("\", "/")

        Using SW As New StreamWriter( _
            Path.Combine( _
                AppPath, _
                "log.txt" _
            ), _
            True _
        )

            Try
                Log(SW, EnvironmentCommandLine)

                Dim matches As List(Of String) = rxMatch(rxProg, EnvironmentCommandLine.Replace(":\\", "://"))
                If matches.Count = 0 Then _
                    Throw New Exception( _
                        String.Format( _
                            "Malformed URL.", _
                            "" _
                        ) _
                    )

                Select Case matches(0).ToLower
                    Case "source"
                        StartProcess(Path.Combine(AppPath, "source\pixie.sln"))

                    Case "schema"
                        StartProcess(Path.Combine(AppPath, "valid.xsd"))

                    Case "config"
                        StartProcess(Path.Combine(UserFolder, "programs.xml"))

                    Case "log"
                        StartProcess("notepad", Path.Combine(AppPath, "log.txt"))

                    Case "help"
                        StartProcess(Path.Combine(AppPath, "help.html"))

                    Case Else

                        Dim progNode As XmlNode

                        With doc
                            ' Load config
                            Try
                                .Load( _
                                    XmlReader.Create( _
                                        IO.Path.Combine( _
                                            UserFolder, _
                                            "programs.xml" _
                                        ), _
                                        Settings _
                                    ) _
                                )

                            Catch EX As Exception
                                Throw New Exception( _
                                    String.Format( _
                                        "Unable to load configuration file. {0}{1}", _
                                        vbCrLf, _
                                        EX.Message _
                                    ) _
                                )

                            End Try

                            ' Parameter validation
                            progNode = doc.SelectSingleNode( _
                                String.Format( _
                                    "pixie/program[@name='{0}']", _
                                    matches(0) _
                                ) _
                            )
                            ValidProg(matches(0), progNode)

                        End With

                        ' Start the process
                        StartProcess(progNode)

                End Select

            Catch ex As Exception
                MsgBox(ex.Message)
                Log(SW, ex.Message)
                Log(SW, ex.StackTrace)

            End Try

        End Using

    End Sub

#End Region

End Module
