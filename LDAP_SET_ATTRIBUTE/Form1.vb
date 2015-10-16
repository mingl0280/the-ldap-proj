Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Text.Encoding


Public Class Form1

    Private DEntry As DirectoryEntry = New DirectoryEntry("LDAP://LocalHost", "Administrator", "asdf4321!!", AuthenticationTypes.Secure)
    Private TimeCode_Staff As String = "0 0 0 0 192 255 0 192 255 0 192 255 0 192 255 0 192 255 0 0 0"
    Private TimeCode_Executive As String = "255 31 248 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255 255"
    Private TimeCode_Manager As String = "3 0 0 0 192 255 3 192 255 3 192 255 3 192 255 3 192 255 3 192 255"


    ''' <summary>
    ''' Nothing important...
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text += "Loading..."
    End Sub

    ''' <summary>
    ''' The sub that calls the program to set Logon Time for each user
    ''' </summary>
    ''' <param name="input">AD Filter String</param>
    ''' <remarks></remarks>
    Sub BeginCall(ByVal input As String)
        Dim testdel = getDirectoryEntry(input)

        For Each elem In testdel
            Dim suc = SetUserDtAttribute(elem)
            If suc Then TextBox1.Text += "--->Successful" Else TextBox1.Text += "--->Failed."
            TextBox1.Text += vbCrLf
        Next
    End Sub

    ''' <summary>
    ''' The sub that calls the program to set Home Drive and Home Directory for each user
    ''' </summary>
    ''' <param name="input">AD Filter String</param>
    ''' <remarks></remarks>
    Sub CSetHomeDrive(ByVal input As String)
        Dim testdel = getDirectoryEntry(input)

        For Each elem In testdel
            Dim suc = SetUserHDrAttribute(elem)
            If suc Then TextBox1.Text += "--->Successful" Else TextBox1.Text += "--->Failed."
            TextBox1.Text += vbCrLf
        Next
    End Sub

    ''' <summary>
    ''' Convert a Octect String into byte array
    ''' </summary>
    ''' <param name="input"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function OctectStringToBytes(ByVal input As String) As Byte()
        Dim strarr() As String = input.Split(" ")
        Dim byteArr() As Byte
        ReDim byteArr(UBound(strarr))
        For i = 0 To UBound(strarr)
            byteArr(i) = CByte(strarr(i))
        Next
        Return byteArr
    End Function

    ''' <summary>
    ''' Set the LogonTime for specified user
    ''' </summary>
    ''' <param name="din">the DirectoryEntry for input, which is a record of the Account</param>
    ''' <returns>The Operation is successful or not</returns>
    ''' <remarks></remarks>
    Private Function SetUserDtAttribute(ByRef din As DirectoryEntry) As Boolean
        Try
            Dim MoF As String = ""
            MoF = din.Properties("memberOf").Value
            If MoF.Contains("StaffMembers") Then
                TextBox1.Text += din.Name + " is a STAFF MEMBER, SETTING LOGIN TIME: 7AM – 5PM on Monday – Friday."
                SetLogonHours(din, TimeCode_Staff)
            ElseIf MoF.Contains("Manager") Then
                TextBox1.Text += din.Name + " is a MANAGER, SETTING LOGIN TIME: Monday – Saturday from 7AM – 7PM."
                SetLogonHours(din, TimeCode_Manager)
            ElseIf MoF.Contains("Executive") Then
                TextBox1.Text += din.Name + " is an EXECUTIVE MEMBER, SETTING LOGIN TIME: 24/7, but are not allowed to on Sunday from 6AM – 12PM."
                SetLogonHours(din, TimeCode_Executive)
            End If
            din.CommitChanges()
            Return True
        Catch ex As Exception
            TextBox1.Text += ex.Message + ex.Source + ex.StackTrace
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Set the HomeDrive and HomeDirectory for specified user
    ''' </summary>
    ''' <param name="din">the DirectoryEntry for input, which is a record of the Account</param>
    ''' <returns>The Operation is successful or not</returns>
    ''' <remarks></remarks>
    Private Function SetUserHDrAttribute(ByRef din As DirectoryEntry) As Boolean
        Try
            Dim MoF As String = ""
            MoF = din.Properties("memberOf").Value
            If MoF.Contains("StaffMembers") Or MoF.Contains("Manager") Then
                TextBox1.Text += din.Name + " is a STAFF MEMBER, SETTING Home Drive S"
                SetHomeDriveProperty(din, "S")
            ElseIf MoF.Contains("Executive") Then
                TextBox1.Text += din.Name + " is an EXECUTIVE MEMBER, SETTING Home Drive O"
                SetHomeDriveProperty(din, "O")
            End If
            din.CommitChanges()
            Return True
        Catch ex As Exception
            TextBox1.Text += ex.Message + ex.Source + ex.StackTrace
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Change the LogonHour for specified DirectoryEntry
    ''' </summary>
    ''' <param name="ent">The Entry that needs to be changed</param>
    ''' <param name="TimeValue">The Octect String Value</param>
    ''' <remarks></remarks>
    Private Sub SetLogonHours(ByRef ent As DirectoryEntry, ByVal TimeValue As String)
        If ent.Properties.Contains("logonHours") Then
            ent.Properties("logonHours").Value = OctectStringToBytes(TimeValue)
        Else
            ent.Properties("logonHours").Add(OctectStringToBytes(TimeValue))
        End If
    End Sub

    ''' <summary>
    ''' Change the HomeDrive and HomeDirectory value for a DirectoryEntry
    ''' </summary>
    ''' <param name="ent">The Entry that needs to be changed</param>
    ''' <param name="letter">The HomeDrive letter desiered to set for the Entry</param>
    ''' <remarks></remarks>
    Private Sub SetHomeDriveProperty(ByRef ent As DirectoryEntry, ByVal letter As String)
        If ent.Properties.Contains("homeDrive") Then
            ent.InvokeSet("homeDrive", letter)
        Else
            ent.InvokeSet("homeDrive", letter)
        End If
        If ent.Properties.Contains("homeDirectory") Then
            ent.InvokeSet("homeDirectory", "\\svr.capsuleco.com\CapsuleHome\" + ent.Properties("sAMAccountName").Value)
        Else
            ent.InvokeSet("homeDirectory", "\\svr.capsuleco.com\CapsuleHome\" + ent.Properties("sAMAccountName").Value)
        End If

    End Sub

    ''' <summary>
    ''' Get the list of DirectoryEntries matches Filter
    ''' </summary>
    ''' <param name="FilterString">AD Filter String(I set as all users)</param>
    ''' <returns>List of DirectoryEntry</returns>
    ''' <remarks>In this program, A DirectoryEntry represents an Account record</remarks>
    Private Function getDirectoryEntry(ByVal FilterString As String) As List(Of DirectoryEntry)
        Dim DSearcher As DirectorySearcher = New DirectorySearcher(DEntry)
        TextBox1.Text += "ADS Load" + vbCrLf 'for debug
        DSearcher.Filter = FilterString
        TextBox1.Text += "Filter Filted" + vbCrLf 'for debug
        DSearcher.SearchScope = SearchScope.Subtree
        Dim Sresult() As SearchResult
        Dim SResultCollection As SearchResultCollection = DSearcher.FindAll()
        ReDim Sresult(SResultCollection.Count - 1)
        TextBox1.Text += "search over" + vbCrLf 'for debug
        SResultCollection.CopyTo(Sresult, 0)
        Dim RetDEList As New List(Of DirectoryEntry)
        TextBox1.Text += SResultCollection.Count.ToString + vbCrLf
        For i As Integer = 0 To UBound(Sresult)
            RetDEList.Add(New DirectoryEntry(Sresult(i).Path))
        Next
        Return RetDEList
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        BeginCall(TextBox2.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CSetHomeDrive(TextBox2.Text)
    End Sub
End Class
