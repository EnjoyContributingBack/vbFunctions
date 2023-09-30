Imports Microsoft.Office.Interop

Public Class msWordLink
    Private msApp As Word.Application
    Private msDoc As Word.Document

    Private currenWordProcessId As Integer

    Private Function getProcessID() As Integer
        Dim pId As Integer = 0

        For Each proc As Process In Process.GetProcessesByName("WINWORD")
            pId += proc.Id
        Next proc

        Return pId
    End Function

    Public Function IsMsWordLinked(ByVal fileName As String) As Boolean
        Try
            Dim oPid As Integer = getProcessID()

            msApp = CreateObject("Word.Application")
            msDoc = msApp.Documents.Open(fileName)

            Dim nPid As Integer = getProcessID()
            currenWordProcessId = nPid - oPid

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub findReplace(ByVal findTxt() As String, ByVal replaceTxt() As String, ByVal replacetimes() As Byte, ByVal itmcount As Byte)
        Dim dCont As Object

        dCont = msDoc.Content

        Dim replacestring As String = String.Empty, strReplace As String = String.Empty
        Dim maxLen As Integer = 251
        Dim n As Integer = findTxt.Length - 1
        For i As Integer = 0 To n
            Try
                'replacetimes = 1 for replace one occurrence, 2 for replace all.
                replacestring = replaceTxt(i)
                For j As Integer = 0 To itmcount - 1
                    If (Not replacestring Is Nothing) Then
                        If (replacestring.Length > maxLen) Then
                            strReplace = replacestring.Substring(0, maxLen)
                            replacestring = replacestring.Substring(maxLen, replacestring.Length - maxLen)
                        Else
                            strReplace = replacestring
                            replacestring = String.Empty
                        End If
                    End If
                    dCont.Find.Execute(findTxt(i), , True, , , , , , , strReplace, replacetimes(i))
                    dCont = msDoc.Content 'Refresh the pointer.
                Next j
            Catch
                'do nothing.
            End Try
        Next i
    End Sub

    Public Sub saveDoc()
        msDoc.Save()
        Try
            msApp.Quit()
            If currenWordProcessId > 0 Then
                Dim proc As Process = Process.GetProcessById(currenWordProcessId)
                proc.Kill()
            End If
        Catch ex As Exception
            'do nothing.
        End Try
    End Sub

    Public Sub showDoc()
        msApp.Visible = True
    End Sub

    Protected Overrides Sub Finalize()
        Try
            'if the normal process does not kill the process then second time process killing
            msDoc = Nothing
            msApp = Nothing
        Catch ex As Exception
            'do nothing.
        End Try

        MyBase.Finalize()
    End Sub
End Class
