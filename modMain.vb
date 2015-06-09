Option Strict On

' This program shows the shared folders on the local computer, or on a remote computer
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started June 9, 2015

' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/ or http://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0

Imports System.Collections.Generic
Imports System.IO


Module modMain

    Public Const PROGRAM_DATE As String = "June 9, 2015"

    Private Structure udtShareInfo
        Public Name As String
        Public Path As String
        Public Caption As String
    End Structure

    Public Sub Main()

        Dim objParseCommandLine As New clsParseCommandLine

        Try
            objParseCommandLine.ParseCommandLine()

            If objParseCommandLine.IsParameterPresent("?") OrElse
               objParseCommandLine.IsParameterPresent("h") OrElse
               objParseCommandLine.IsParameterPresent("help") Then

                ShowProgramHelp()

                Exit Sub

            End If

            Dim showAdminShares = objParseCommandLine.IsParameterPresent("Admin")

            If objParseCommandLine.NonSwitchParameterCount > 0 Then
                Dim hostName = objParseCommandLine.RetrieveNonSwitchParameter(0)
                ListWindowsShares(hostName, showAdminShares)
            Else
                ListWindowsShares(showAdminShares)
            End If

        Catch ex As Exception
            ShowErrorMessage("Error occurred in modMain->Main: " & Environment.NewLine & ex.Message)
        End Try
      
    End Sub

    Private Function GetAppPath() As String
        Return Reflection.Assembly.GetExecutingAssembly().Location
    End Function

    Private Function GetAppVersion() As String
        'Return Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")"

        Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString & " (" & PROGRAM_DATE & ")"
    End Function


    Private Sub ListWindowsShares(showAdminShares As Boolean)
        Dim hostName = Net.Dns.GetHostName()

        ListWindowsShares(hostName, showAdminShares)
    End Sub

    Private Sub ListWindowsShares(hostName As String, showAdminShares As Boolean)

        If hostName.StartsWith("\\") Then
            hostName = hostName.Substring(2)
        End If

        If String.IsNullOrEmpty(hostName) Then
            ShowErrorMessage("hostName cannot be empty")
            Return
        End If

        Dim shareList = New List(Of udtShareInfo)
        Dim adminSharesSkipped As Integer

        Dim maxNameLength = 15
        Dim maxPathLength = 15
        Dim queryString = "??"

        Try
            queryString = "\\" & hostName & "\root\cimv2"
            Using shares As New Management.ManagementClass(queryString, "Win32_Share", New Management.ObjectGetOptions())

                For Each share As Management.ManagementObject In shares.GetInstances()
                    Dim shareInfo = New udtShareInfo

                    shareInfo.Name = share("Name").ToString()

                    If Not showAdminShares AndAlso shareInfo.Name.EndsWith("$") Then
                        adminSharesSkipped += 1
                        Continue For
                    End If

                    shareInfo.Caption = share("Caption").ToString()
                    shareInfo.Path = share("Path").ToString()

                    maxNameLength = Math.Max(shareInfo.Name.Length, maxNameLength)
                    maxPathLength = Math.Max(shareInfo.Path.Length, maxPathLength)

                    shareList.Add(shareInfo)
                Next
            End Using

        Catch ex As Exception
            ShowErrorMessage("Error querying for shares using: " & queryString & Environment.NewLine & ex.Message)
            Return
        End Try

        If shareList.Count = 0 Then
            If showAdminShares Then
                Console.WriteLine("Host " & hostName & " has no shared folders or admin shares")
            Else
                Console.WriteLine("Host " & hostName & " has no shared folders")
                Console.WriteLine("(" & adminSharesSkipped & " admin shares; see them with /admin)")
            End If
        Else
            maxNameLength += 2
            maxPathLength += 2

            Console.WriteLine("Name".PadRight(maxNameLength) & "Path".PadRight(maxPathLength) & "Caption")

            For Each share In shareList
                Console.WriteLine(share.Name.PadRight(maxNameLength) & share.Path.PadRight(maxPathLength) & share.Caption)
            Next

            If adminSharesSkipped > 0 Then
                Console.WriteLine()
                Console.WriteLine(adminSharesSkipped & " hidden admin shares not listed")
                Console.WriteLine("(to see them, use /admin)")
            End If

        End If

    End Sub

    Private Sub ShowErrorMessage(ByVal strMessage As String)
        Dim strSeparator As String = "------------------------------------------------------------------------------"

        Console.WriteLine()
        Console.WriteLine(strSeparator)
        Console.WriteLine(strMessage)
        Console.WriteLine(strSeparator)
        Console.WriteLine()

    End Sub

    Private Sub ShowProgramHelp()

        Try
            Console.WriteLine("This program shows the shared folders on the local computer, or on a remote computer")
            Console.WriteLine()

            Console.WriteLine("Program syntax:")
            Console.WriteLine(Path.GetFileName(GetAppPath()) & " [HostName] /Admin")
            Console.WriteLine()
            Console.WriteLine("The host name is optional.  If missing, then shows local shares")
            Console.WriteLine("Use /Admin to see the hidden, administrative shares")

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2010")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://omics.pnl.gov/ or http://panomics.pnnl.gov/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            Threading.Thread.Sleep(750)

        Catch ex As Exception
            Console.WriteLine("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub


End Module


