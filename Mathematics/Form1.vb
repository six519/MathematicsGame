Imports ADODB
Public Class Form1
    'palitan ang constant kung ilang segundo kada tanong ang blang
    Const FIX_BILANG As Integer = 30

    Dim conn As New Connection
    Dim rs As New ADODB.Recordset

    Dim level As Integer = 0
    'from level 1 to 4
    Dim score As Integer = 0

    Dim AnsNum(100) As Integer
    'kelangang ang value lagi ng questnum ay 1 sa umpisa
    Dim QuestNum As Integer = 1
    Dim QuestMax As Integer
    Dim Quest(100) As String
    Dim QuestA(100) As String
    Dim QuestB(100) As String
    Dim QuestC(100) As String
    Dim QuestD(100) As String
    Dim QuestRight(100) As Integer
    'para sa fastest
    Dim Segundo As Integer = 0

    'para sa time pressure
    Dim bilang As Integer = FIX_BILANG

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'open the database
        Try
            'try connecting to database
            conn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & Application.StartupPath & "\db.mdb"
            conn.Open()
        Catch ero1 As System.Runtime.InteropServices.COMException
            MsgBox("Cannot open database", MsgBoxStyle.Critical, "MATHEMATICS")
            End
        End Try
        'disable menu on flash
        shock.Menu = False
        shock.Left = 0
        shock.Top = 0
        'set the movie path
        shock.Movie = Application.StartupPath & "\main.swf"
    End Sub

    Private Sub initQuest()
        On Error GoTo errhand
        Dim tableName As String = ""
        Dim n As Integer = 1

        Select Case level
            Case 1
                tableName = "table_one"
            Case 2
                tableName = "table_two"
            Case 3
                tableName = "table_three"
            Case 4
                tableName = "table_four"
        End Select

        rs.CursorLocation = CursorLocationEnum.adUseClient

        rs.Open("select * from " & tableName, conn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
        rs.MoveLast()
        QuestMax = rs.RecordCount
        rs.MoveFirst()

        'MsgBox(QuestMax)


        Do Until n > QuestMax
            Quest(n) = rs("question").Value
            QuestA(n) = rs("q1").Value
            QuestB(n) = rs("q2").Value
            QuestC(n) = rs("q3").Value
            QuestD(n) = rs("q4").Value
            QuestRight(n) = Val(rs("right_answer").Value)
            n = n + 1
            rs.MoveNext()
        Loop


        rs.Close()
        Exit Sub
errhand:


    End Sub


    Private Sub GiveQ()
        If QuestNum > QuestMax Then
            Timer2.Enabled = False
            bilang = FIX_BILANG
            'stop ang timer
            Timer1.Enabled = False
            'computation ng score
            CompuTan()
            shock.SetVariable("gotoeval", "true")
        Else
            shock.SetVariable("koras", FIX_BILANG)
            bilang = FIX_BILANG

            shock.SetVariable("tanong", Quest(QuestNum) & "<br>" _
            & "A. " & QuestA(QuestNum) & "<br>" & "B. " & QuestB(QuestNum) & "<br>" _
            & "C. " & QuestC(QuestNum) & "<br>" & "D. " & QuestD(QuestNum))
            'MsgBox Quest(3)
            QuestNum = QuestNum + 1
        End If
    End Sub


    Private Sub CompuTan()
        Dim n As Integer = 1
        Dim Ilalag As String = ""

        Do Until n > QuestMax

            'MsgBox(AnsNum(n) & " : " & QuestRight(n))

            If AnsNum(n) = QuestRight(n) Then
                'right
                score = score + 1
            End If

            n = n + 1
        Loop

        If (score * 100 \ QuestMax) < 50 Then
            Ilalag = "You are stupid?Do you know it?" & vbNewLine
        ElseIf (score * 100 \ QuestMax) = 100 Then
            'MsgBox "Wow!Your I.Q. level is equal to Einstein!", vbInformation, "MATHIGH"
            Ilalag = "Wow!Your I.Q. level is equal to Einstein!" & vbNewLine
        ElseIf (score * 100 \ QuestMax) > 50 Then
            'MsgBox "Nice!But not too smart!", vbInformation, "MATHIGH"
            Ilalag = "Nice!But not too smart!" & vbNewLine
        End If

        Ilalag = Ilalag & "Your Score is: " & (score * 10) & vbNewLine & "Your Time is: " & Segundo & vbNewLine & _
        "You got " & score & " correct answer out of " & QuestMax & " questions!"

        ' shock.SetVariable("corek.evalog", Ilalag)
        shock.SetVariable("balyos", Ilalag)
        'MsgBox(Ilalag)
    End Sub


    Private Sub ResetAll()
        QuestNum = 1
        score = 0
        shock.SetVariable("gotoeval", "false")
        level = 0
        Segundo = 0
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Segundo = Segundo + 1
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        bilang = bilang - 1
        shock.SetVariable("koras", bilang)

        If bilang = 0 Then
            'tutunog
            shock.SetVariable("buzz", "true")
            'bgay na uli ng bagong tanong
            'para ng tinawag ang geta command
            AnsNum(QuestNum - 1) = 0 'maling sagot
            GiveQ()
        End If

    End Sub

    Private Sub shock_FSCommand1(ByVal sender As Object, ByVal e As AxShockwaveFlashObjects._IShockwaveFlashEvents_FSCommandEvent) Handles shock.FSCommand
        'tanggalin ito para malaman ang error
        'On Error Resume Next
        Dim scoreStr As String = ""
        Dim x As Integer = 0
        'mga fscommands galing sa flash object
        Select Case e.command
            Case "taymer"
                'activate timer
                Timer1.Enabled = True
                Timer2.Enabled = True
            Case "hscore"
                'view highscore
                rs.Open("select * from table_score order by score desc", conn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                For x = 1 To 5
                    scoreStr = scoreStr & rs("player").Value & " = " & rs("score").Value & "pt/s = " & rs("time").Value & "sec/s" & vbLf
                    rs.MoveNext()
                Next

                shock.SetVariable("topscore", scoreStr)
                rs.Close()

                scoreStr = ""


                rs.Open("select * from table_score order by time asc", conn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)

                For x = 1 To 5
                    scoreStr = scoreStr & rs("player").Value & " = " & rs("score").Value & "pt/s = " & rs("time").Value & "sec/s" & vbLf
                    rs.MoveNext()
                Next

                shock.SetVariable("topscore2", scoreStr)
                rs.Close()

            Case "savescore"
                rs.Open("select * from table_score", conn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                rs.AddNew()
                rs("player").Value = e.args
                rs("score").Value = score * 10
                rs("time").Value = Segundo
                rs.Update()
                ResetAll()
                rs.Close()
            Case "geta"
                AnsNum(QuestNum - 1) = Val(e.args)
                GiveQ()
                'MsgBox("test")
            Case "initq"
                shock.SetVariable("koras", FIX_BILANG)
                'initialization of question
                initQuest()
                'bigay ng unang question
                GiveQ()
            Case "setlevel"
                'set the level number
                level = Val(e.args)
            Case "help"
                rs.Open("select * from table_help", conn, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic)
                shock.SetVariable("helper.help_box", rs("help").Value)
                rs.Close()
            Case "close"
                conn.Close()
                Me.Close()
        End Select
    End Sub
End Class
