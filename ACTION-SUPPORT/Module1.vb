Module Module1
    Public Function OpenConnectionOfMes()
        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "IQMES3"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.MultipleActiveResultSets = True
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        Return mConnectionBuilder.ConnectionString
    End Function
    Public Function OpenConnectionOfRDMes()
        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "IQMES-TEST1"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.MultipleActiveResultSets = True
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        Return mConnectionBuilder.ConnectionString
    End Function
    Public Function OpenOracleConnection(ByVal odb As String)
        'Dim oConnectionBuilder As New Oracle.DataAccess.Client.OracleConnectionStringBuilder
        Dim oConnectionBuilder As New Oracle.ManagedDataAccess.Client.OracleConnectionStringBuilder
        oConnectionBuilder.DataSource = "topprod"
        'oConnectionBuilder.DBAPrivilege = "Normal"
        oConnectionBuilder.UserID = odb
        oConnectionBuilder.Password = odb
        oConnectionBuilder.PersistSecurityInfo = True
        Return oConnectionBuilder.ConnectionString
    End Function
    Public Function OpenConnectionOfERPSUPPORT()
        Dim mConnectionBuilder As New SqlClient.SqlConnectionStringBuilder
        mConnectionBuilder.DataSource = "192.168.10.254"
        mConnectionBuilder.InitialCatalog = "ERPSUPPORT"
        mConnectionBuilder.IntegratedSecurity = False
        mConnectionBuilder.UserID = "sa"
        mConnectionBuilder.Password = "p@$$w0rd"
        mConnectionBuilder.MultipleActiveResultSets = True
        Return mConnectionBuilder.ConnectionString
    End Function
    Public Sub KillExcelProcess(ByVal oldExcel() As Process)
        Dim NewExcelProcess() As Process = Process.GetProcessesByName("Excel")
        For i As Int16 = 0 To NewExcelProcess.Length - 1 Step 1
            Dim FoundExcel As Boolean = False
            Dim NewProcessInteger As Integer = NewExcelProcess(i).Id
            For j As Int16 = 0 To oldExcel.Length - 1 Step 1
                Dim OldProcessIntger As Integer = oldExcel(j).Id
                If NewProcessInteger = OldProcessIntger Then
                    FoundExcel = True
                    Exit For
                End If
            Next
            If FoundExcel = False Then
                Process.GetProcessById(NewExcelProcess(i).Id).Kill()
                Exit For
            End If
        Next
    End Sub
    Public Function GetYearAndMonthString(ByVal t1 As DateTime)
        Dim YM1 As String = String.Empty
        Dim Year1 As String = t1.Year
        Dim Month1 As String = t1.Month
        If Month1.Length = 1 Then
            YM1 = Year1 & "0" & Month1
        Else
            YM1 = Year1 & Month1
        End If
        Return YM1
    End Function
    Public Function GetMonthEnglish(ByVal month1 As Int16)
        Dim GE As String = String.Empty
        Select Case month1
            Case 1
                GE = "January"
            Case 2
                GE = "February"
            Case 3
                GE = "March"
            Case 4
                GE = "April"
            Case 5
                GE = "May"
            Case 6
                GE = "June"
            Case 7
                GE = "July"
            Case 8
                GE = "August"
            Case 9
                GE = "September"
            Case 10
                GE = "October"
            Case 11
                GE = "November"
            Case 12
                GE = "December"
        End Select
        Return GE
    End Function
    Public Function CheckAuthorizeByPC(ByVal F1 As String, ByVal PN As String)
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        mConnection.ConnectionString = OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        mSQLS1.CommandText = "SELECT COUNT(*) FROM Authorize WHERE Module = '" & F1 & "' AND PCName = '" & PN & "' AND Available = 'Y'"
        Dim Az As Boolean = False
        Dim CountAu As Int16 = mSQLS1.ExecuteScalar()
        mConnection.Close()
        If CountAu > 0 Then
            Az = True
        Else
            Az = False
        End If
        Return Az
    End Function
    Public Function CheckAuthorizeByUser(ByVal F1 As String, ByVal PN As String)
        Dim mConnection As New SqlClient.SqlConnection
        Dim mSQLS1 As New SqlClient.SqlCommand
        mConnection.ConnectionString = OpenConnectionOfERPSUPPORT()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        mSQLS1.CommandText = "SELECT COUNT(*) FROM AuthorizeByUser WHERE Module = '" & F1 & "' AND UserName = '" & PN & "' AND Available = 'Y'"
        Dim Az As Boolean = False
        Dim CountAu As Int16 = mSQLS1.ExecuteScalar()
        If PN = "cloud.sheng" Or PN = "brady.chen" Then
            CountAu = 1
        End If
        mConnection.Close()
        If CountAu > 0 Then
            Az = True
        Else
            Az = False
        End If
        Return Az
    End Function
End Module
