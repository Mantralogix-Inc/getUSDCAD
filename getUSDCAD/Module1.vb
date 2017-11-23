Imports System.IO
Imports System.Net
Imports System.Windows.Forms
Imports System.Data.SqlClient

Module Module1
    Public CONN As String
    Dim msDatabase As String
    Dim msTable As String
    Dim msGetDate As String
    Sub Main()
        getSqlIds()
        ' **********************************************************************************
        ' Bank of Canada updates rate at 16:30 ET time
        ' Get yesterday date
        ' api.fixer.io only allows limited checks. Using Bank of Canada
        ' **********************************************************************************

        Dim lsApi As String = ""
        '"http://api.fixer.io/latest?base=CAD&symbols=USD"
        lsApi = "https://www.bankofcanada.ca/valet/observations/FXCADUSD/json?start_date=" & msGetDate & "&end_date=" & msGetDate

        Dim lsFrom As String = ""
        Dim lsTo As String = ""

        Dim responseFromServer As String = ""
        Try
            Dim request As WebRequest =
                      WebRequest.Create(lsApi)
            Dim response As WebResponse = request.GetResponse()
            Dim dataStream As Stream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            responseFromServer = reader.ReadToEnd()
            reader.Close()
            response.Close()
        Catch ex As WebException
            ' There was an error, write down the cause in the log
            writelog("Conversion.txt", "Problem:" & ex.Message & vbCrLf, True)
        Finally
        End Try
        If responseFromServer.Trim <> "" Then
            Dim lsRate As String = fnGetBetween(responseFromServer, ":{""v"":", "}}")
            Dim lsDate As String = fnGetBetween(responseFromServer, "{""d"":""", """,""")
            Dim lsSql As String = My.Resources.insertOrUpdateConvertion
            Dim REVCOURS As Double = 0
            If IsNumeric(lsRate) Then
                REVCOURS = Math.Round(1 / Val(lsRate), 4)
                lsSql = Replace(lsSql, "<<DATABASE>>", msDatabase)
                lsSql = Replace(lsSql, "<<TABLE>>", msTable)
                lsSql = Replace(lsSql, "<<CURDEN>>", "CAD")
                lsSql = Replace(lsSql, "<<CUR>>", "USD")
                lsSql = Replace(lsSql, "<<CHGSTRDAT>>", lsDate)
                lsSql = Replace(lsSql, "<<CHGRAT>>", lsRate)
                lsSql = Replace(lsSql, "<<REVCOURS>>", REVCOURS.ToString)
                sqlExe(lsSql)
                ' Write down the rate and date in the log
                writelog("Conversion.txt", "Ran on " & Now.ToString & "   downloaded Date: " & lsDate & " >> CAD=" & lsRate & vbCrLf, True)
            Else
                ' No data for the date. Write down the problem in the log
                writelog("Conversion.txt", "Ran on " & Now.ToString & "  no data available for " & msGetDate & vbCrLf, True)
            End If
        End If
    End Sub
    Public Sub sqlExe(mySql As String)
        ' ***************************************************
        ' Insert a new item to temp table or delete. Use any sql
        ' ***************************************************

        Dim myConn As SqlConnection = New SqlConnection(CONN)
        Dim myCommand As SqlCommand = New SqlCommand(mySql, myConn)
        Try
            myConn.Open()
            myCommand.CommandTimeout = 360
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            Debug.Print(mySql)
            myErrorMsg(ex, System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf & mySql)
            'MsgBox(ex.InnerException.ToString & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Error on Sub/Function:  " & System.Reflection.MethodInfo.GetCurrentMethod.Name & vbCrLf & vbCrLf & mySql)
            Using writer As StreamWriter = New StreamWriter(Application.StartupPath & "\sqlExe.txt")
                writer.Write(mySql)
            End Using
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
            myCommand.Dispose()
            myConn.Dispose()
        End Try
    End Sub
    Private Sub getSqlIds()
        Dim lsID As String
        Dim lsPW As String
        Dim arguments As String() = Environment.GetCommandLineArgs()
        ' msDatabase = "sagex3"
        ' msTable = "KNC01"
        Dim lsDataSource As String
        ' lsDataSource = "172.16.1.23\SAGEX3"
        msGetDate = Format(Now, "yyyy-MM-dd")

        For x As Integer = 0 To arguments.Count - 1
            If arguments(x).Contains("T:") Then
                msTable = Replace(arguments(x), "T:", "")
            End If
            If arguments(x).Contains("D:") Then
                msDatabase = Replace(arguments(x), "D:", "")
            End If
            If arguments(x).ToUpper.Contains("S:") Then
                lsDataSource = Replace(arguments(x).ToUpper, "S:", "")
            End If
            If arguments(x).ToUpper.Contains("Y:TRUE") Then
                ' Run it for yesterday
                msGetDate = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd")
            End If
        Next


        lsID = msTable
        lsPW = "tiger"
        CONN = _
       "uid=" & lsID & ";pwd=" & lsPW & ";Persist Security Info=False;" + _
       "Initial Catalog=" & msDatabase & ";Data Source=" & lsDataSource & ";Connection Timeout=300"

    End Sub
    Function fnGetBetween(ByVal sOriginal As String, ByVal sFrom As String, ByVal sTo As String) As String
        Dim lsTemp = Split(sOriginal & sFrom, sFrom, -1)
        Dim lsTemp2 = Split(lsTemp(1), sTo, -1)
        Return lsTemp2(0)
    End Function
    Public Sub myErrorMsg(ByVal myex As Exception, ByVal mysub As String)
        Dim lsBuildError As String = ""
        If myex.InnerException IsNot Nothing Then
            lsBuildError = myex.InnerException.Message.ToString & vbCrLf
        End If
        MsgBox(lsBuildError & myex.Message & vbCrLf & vbCrLf & "Error on Sub/Function:  " & mysub & vbCrLf & vbCrLf & myex.StackTrace)
    End Sub
    Public Sub writelog(ByVal myFileName As String, ByVal log As String, ByVal Append As Boolean)
        Using writer As StreamWriter = New StreamWriter(Application.StartupPath & "\" & myFileName, Append)
            writer.Write(log)
        End Using
    End Sub
End Module
