Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory


Public Class Form1

    Private DEntry As DirectoryEntry = New DirectoryEntry("LDAP://LocalHost", "Administrator", "asdf4321!!", AuthenticationTypes.Secure)
    Private TimeCode_Staff As String = "0 0 0 0 192 255 0 192 255 0 192 255 0 192 255 0 192 255 0 0 0"
    Private TimeCode_Executive As String = "255 31 248 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255"
    Private TimeCode_Manager As String = "3 0 0 0 192 255 3 192 255 3 192 255 3 192 255 3 192 255 3 192 255"



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text += "Loading..."
        Dim testdel = getDirectoryEntry("(&(objectClass=user))")

        For Each elem In testdel
            Dim suc = SetUserDtAttribute(elem)
            If suc Then TextBox1.Text += "--->Successful" Else TextBox1.Text += "--->Failed."
            TextBox1.Text += vbCrLf
        Next

    End Sub

    Private Function SetUserDtAttribute(ByRef lin As DirectoryEntry) As Boolean
        Try
            Dim MoF As String = ""
            MoF = lin.Properties("memberOf").Value
            If MoF.Contains("StaffMembers") Then
                TextBox1.Text += lin.Name + " is a STAFF MEMBER, SETTING LOGIN TIME: 7AM – 5PM on Monday – Friday."
                SetLogonHours(lin, TimeCode_Staff)
                lin.InvokeSet("logonHours", DirectCast(TimeCode_Staff, Object))
            ElseIf MoF.Contains("Manager") Then
                TextBox1.Text += lin.Name + " is a MANAGER, SETTING LOGIN TIME: Monday – Saturday from 7AM – 7PM."
                SetLogonHours(lin, TimeCode_Manager)
            ElseIf MoF.Contains("Executive") Then
                TextBox1.Text += lin.Name + " is an EXECUTIVE MEMBER, SETTING LOGIN TIME: 24/7, but are not allowed to on Sunday from 6AM – 12PM."
                SetLogonHours(lin, TimeCode_Executive)
            End If
            lin.CommitChanges()
            Return True
        Catch ex As Exception
            TextBox1.Text += ex.Message + ex.Source + ex.StackTrace
            Return False
        End Try
    End Function

    Private Sub SetLogonHours(ByRef ent As DirectoryEntry, ByVal TimeValue As String)
        If ent.Properties.Contains("logonHours") Then
            ent.Properties("logonHours").Value = TimeValue
        Else
            ent.Properties("logonHours").Add(TimeValue)
        End If
    End Sub

    Private Function getDirectoryEntry(ByVal FilterString As String) As List(Of DirectoryEntry)
        Dim DSearcher As DirectorySearcher = New DirectorySearcher(DEntry)
        TextBox1.Text += "ADS Load" + vbCrLf
        DSearcher.Filter = FilterString
        TextBox1.Text += "Filter Filted" + vbCrLf
        DSearcher.SearchScope = SearchScope.Subtree
        Dim Sresult() As SearchResult
        Dim SResultCollection As SearchResultCollection = DSearcher.FindAll()
        ReDim Sresult(SResultCollection.Count - 1)
        TextBox1.Text += "search over" + vbCrLf
        SResultCollection.CopyTo(Sresult, 0)
        Dim RetDEList As New List(Of DirectoryEntry)
        TextBox1.Text += SResultCollection.Count.ToString + vbCrLf
        For i As Integer = 0 To UBound(Sresult)
            RetDEList.Add(New DirectoryEntry(Sresult(i).Path))
            'TextBox1.Text += Sresult(i).Path + vbCrLf
        Next
        Return RetDEList
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        getDirectoryEntry(TextBox2.Text)
    End Sub
End Class
