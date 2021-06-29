Imports System.Threading
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Public Class Form1

#Region "GlobalVariableInit"


    'Initialize Vairables
    Dim DataFile As String = ""
    Dim SaveFile As String = ""
    Dim GroupFile As String = ""
    Dim ViewFile As String = ""
    Private _LastOffset As Long
    Private _LastOffset2 As Long
    Private _LastOffset3 As Long
    Private _LastOffset4 As Long
    Private _LastOffset5 As Long
    Private _LastOffset6 As Long
    Private _LastOffset7 As Long
    Private _LastOffset8 As Long
    Private _LastOffset9 As Long

    Dim DelimiterCount As Integer = 64
    Dim SubDimensionsCount As Integer = 2
    Dim LastProcessedName As String = ""
    Dim UniqueRacerNamesFound As Integer = 0
    Dim RacerNameTimesFound As Integer = 0
    Dim CurrentRacerNumber As Integer = 0

    Dim MaxGroupSize As Integer = 8
    Dim CurrentGroup As Integer = 1
    Dim CurrentGroupMembers As Integer = 0
    Dim TotalGroups As Integer = 0

    Dim Racers(10, 5, 1) As String
    Dim RaceHistory(10, 1) As String
    Dim OutputFileArray(10, 1) As String

    Dim RaceDataArraySize As Integer = 250
    Dim RacerList(RaceDataArraySize) As String
    Dim RacerMatrix(RaceDataArraySize, RaceDataArraySize, 8, 2) As Single
    Dim RacerCache(10, 12) As String
    Dim CacheInc As Integer = 1
    Dim MatrixCalcMode As Integer = 0

    Dim HistoryInc As Integer = 1
    Dim OutputDataFormat As Integer = 0
    Dim CollisionSubmittedNames As Integer = 0
    Dim CollisionX As Integer = 0
    Dim CollisionY As Integer = 0

    Dim RecomputeIDs As Boolean = False
    Dim AbortCollision As Integer = 0

    Dim LinesInFile As Integer = 0
    Dim CurrentLineNumber As Integer = 0

    Dim BatchRacesProcessed As Integer = 0

    '########## GUI/MISC VARS ##########
    Dim Go As Boolean
    Dim LeftSet As Boolean
    Dim TopSet As Boolean
    Dim HoldLeft As Integer
    Dim HoldTop As Integer
    Dim OffLeft As Integer
    Dim OffTop As Integer
    Dim HoldWidth As Integer
    Dim HoldHeight As Integer

    Dim OrigAlpha As Double = 1

    Dim ImageSelectionMode As Integer = 0

    Dim ImageSelectX1 As Integer = 0
    Dim ImageSelectY1 As Integer = 0

    Dim ImageSelectX2 As Integer = 0
    Dim ImageSelectY2 As Integer = 0


    '########## SCRAPER VARS ##########

    Dim ProgramDir As String = "C:\Users\" & Environment.UserName & "\Desktop\ForeRunner\"
    Dim ProgramArchiveDir As String = "C:\Users\" & Environment.UserName & "\Desktop\ForeRunner\Archive\"
    Dim ProgramCfgDir As String = "C:\Users\" & Environment.UserName & "\Desktop\ForeRunner\CFG\"
    Dim ProgramLogDir As String = "C:\Users\" & Environment.UserName & "\Desktop\ForeRunner\Logs\"
    Dim ProgramScrapeDir As String = "C:\Users\" & Environment.UserName & "\Desktop\ForeRunner\Scrape\"
    Private _ScrapeLastOffset As Long
    Private _ScrapeLastOffset2 As Long
    Private _ScrapeLastOffset3 As Long

    Dim ArchiveFiles As Boolean = True
    Dim PagesArchived As Integer = 0
    Dim PagesScanned As Integer = 0
    Dim ErrorCount As Integer = 0
    Dim ErrorCountMax As Integer = 0
    Dim EngineActive As Integer = 0
    Dim EnginePaused As Integer = 0
    Dim EngineStopped As Integer = 0
    Dim ConcatMaxLines As Integer = 30

    'Vars used for BenchMarking purposes - Not used unless BenchMark is Active in UI settings
    Dim BenchmarkingMode As Integer = 0
    Dim BenchmarkDoEvents As Boolean = False

    Dim BenchmarkPrevTime As Long

    Dim BenchmarkLastIterationTime As Long
    Dim BenchmarkCurrTime As Long
    Dim BenchmarkCurrentIteration As Integer = 0
    Dim BenchmarkCurrentAvgIteration As Integer = 0

    Dim BenchmarkIterationData(10) As Long
    Dim BenchmarkIterationAverages(10) As Long
    Dim BenchmarkStageAverages(10) As Long

    Dim BenchmarkNextUpdateTime As Long
    Dim BenchmarkIterationsThisSecond As Integer = 0


    Dim BenchmarkFinished As Integer = 0


    Public Property ListBox1 As Object


#End Region

    '########## UI Components ##########

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Dim UpperLeft As Point() = {New Point(0, 0)}
        Me.Location = UpperLeft(0)
    End Sub

    Private Sub LoadDataButton_Click(sender As Object, e As EventArgs) Handles LoadDataButton.Click
        OpenFileDialog1.ShowDialog()
        DataFile = OpenFileDialog1.FileName()
        LoadFileTextBox.Text = DataFile
    End Sub

    Private Sub SaveDataButton_Click(sender As Object, e As EventArgs) Handles SaveDataButton.Click
        SaveFileDialog1.ShowDialog()
        SaveFile = SaveFileDialog1.FileName()
        SaveFileTextBox.Text = SaveFile
    End Sub

    Private Sub SaveResultsButton_Click(sender As Object, e As EventArgs) Handles SaveResultsButton.Click
        WriteFile()
    End Sub

    Private Sub SaveResults2Button_Click(sender As Object, e As EventArgs) Handles SaveResults2Button.Click
        WriteFile2()
    End Sub

    Private Sub NameListFileButton_Click(sender As Object, e As EventArgs) Handles NameListFileButton.Click
        OpenFileDialog1.ShowDialog()
        GroupFile = OpenFileDialog1.FileName()
        NameListFileTextBox.Text = GroupFile
    End Sub

    Private Sub NewEventMsg(ByVal input As String)
        EventBox.AppendText(Environment.NewLine & input)
        EventBox.Select(EventBox.TextLength, 0)
        EventBox.ScrollToCaret()
    End Sub

    Private Sub NewEventMsg2(ByVal input As String)
        EventBox2.AppendText(Environment.NewLine & input)
        EventBox2.Select(EventBox2.TextLength, 0)
        EventBox2.ScrollToCaret()
    End Sub

    Private Sub NewEventMsg3(ByVal input As String)
        EventBox3.AppendText(Environment.NewLine & input)
        EventBox3.Select(EventBox3.TextLength, 0)
        EventBox3.ScrollToCaret()
    End Sub

    Private Sub NewEventMsg4(ByVal input As String)
        EventBox4.AppendText(Environment.NewLine & input)
        EventBox4.Select(EventBox4.TextLength, 0)
        EventBox4.ScrollToCaret()
    End Sub

    Public Sub ExecAlg1_Click(sender As Object, e As EventArgs) Handles ExecAlg1.Click
        NewEventMsg("Computing next day's races...")
        CountUniqueRacerOccurrences(DataFile)
        ReDim Racers(UniqueRacerNamesFound, DelimiterCount, SubDimensionsCount)
        ReadDataLines(DataFile)
        If File.Exists(GroupFile) Then
            ReadGroupingFile(GroupFile)
            CalculateRanksInGroup()
        End If
    End Sub

    Private Sub ExecAlg2_Click(sender As Object, e As EventArgs) Handles ExecAlg2.Click
        NewEventMsg("Doing Back-Test, this may take a while...")
        CountLinesInFile(DataFile)
        ReDim Racers(UniqueRacerNamesFound, DelimiterCount, SubDimensionsCount + 1)
        ReDim RaceHistory(LinesInFile, 16)
        ImportBackTestData(DataFile)
        NewEventMsg("Data file imported successfully.")
        NewEventMsg("Calculating time-relative statistics...")
        ComputeBackTestData()
        NewEventMsg("FINISHED: Backtesting finished, results ready to save.")
    End Sub


    '########## Data Parsing Functions ##########

    Private Sub WriteFile()
        Using sw As StreamWriter = File.CreateText(SaveFile)
            For i = 1 To UniqueRacerNamesFound
                sw.WriteLine(CStr(Racers(i, 0, 0)) & "," & CStr(CDbl(Racers(i, 0, 1))) & "," & CStr(Math.Round(CDbl(Racers(i, 1, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 2, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 3, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 4, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 5, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 6, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 7, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 8, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 9, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 10, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 11, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 12, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 13, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 14, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 15, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 16, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 17, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 18, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 19, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 20, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 21, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 22, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 23, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 24, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 25, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 26, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 27, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 28, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 29, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 30, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 31, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 32, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 33, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 34, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 35, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 36, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 37, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 38, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(i, 39, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 40, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 41, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 42, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 43, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 44, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 45, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 46, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 47, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 48, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 49, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 50, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 51, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 52, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 1, 1)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 53, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 54, 0)), 2)) & "," & CStr(Math.Round(CDbl(Racers(i, 55, 2)), 2)))
            Next i
        End Using
        NewEventMsg("SAVED: Saved current program memory array to: " & SaveFile)
    End Sub

    Private Sub WriteFile2()
        Using sw As StreamWriter = File.CreateText(SaveFile)
            For i = 1 To LinesInFile
                Dim NewLineWrite As String = ""
                NewLineWrite = OutputFileArray(i, 0)
                For o = 1 To DelimiterCount
                    NewLineWrite = NewLineWrite & "," & CStr(OutputFileArray(i, o))
                Next o
                sw.WriteLine(NewLineWrite)
            Next i
        End Using
        NewEventMsg("SAVED: Saved current program memory array to: " & SaveFile)
    End Sub

    Private Sub CountUniqueRacerOccurrences(ByVal InputFile As String)
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs2 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr2 As New StreamReader(fs2)
            If _LastOffset < sr2.BaseStream.Length Then
                sr2.BaseStream.Seek(_LastOffset, SeekOrigin.Begin)
                While sr2.Peek() <> -1
                    Dim read As String = sr2.ReadLine()
                    If read.Length > 0 Then
                        CountRacers(read)
                    End If
                End While
                _LastOffset = sr2.BaseStream.Position
            End If
            sr2.Close()
            fs2.Close()
            NewEventMsg("COUNT: " & UniqueRacerNamesFound & " unique racers found and grouped.")
        End If
    End Sub

    Private Sub CountRacers(ByVal InputLine As String)
        Dim RacerName As String = ""
        Dim aa As Integer = 0
        aa = InputLine.IndexOf(",")
        RacerName = InputLine.Remove(0, aa + 1)
        aa = RacerName.IndexOf(",")
        RacerName = RacerName.Remove(aa, RacerName.Length - aa)
        If LastProcessedName = RacerName Then
            'Do Nothing
        Else
            RacerNameTimesFound = 1
            UniqueRacerNamesFound = UniqueRacerNamesFound + 1
            LastProcessedName = RacerName
        End If
    End Sub

    Private Sub ReadDataLines(ByVal InputFile As String)
        Dim DataLineNum As Integer = 0
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs2 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr2 As New StreamReader(fs2)
            If _LastOffset2 < sr2.BaseStream.Length Then
                sr2.BaseStream.Seek(_LastOffset2, SeekOrigin.Begin)
                While sr2.Peek() <> -1
                    Dim read As String = sr2.ReadLine()
                    If read.Length <> 0 Then
                        ProcessDataLines(read)
                    End If
                    DataLineNum = DataLineNum + 1
                End While
                _LastOffset2 = sr2.BaseStream.Position
            End If
            sr2.Close()
            fs2.Close()
            _LastOffset2 = 0
            NewEventMsg("FINISHED: Read " & DataLineNum & " lines, and parsed them into memory.")
        End If
    End Sub

    Private Sub DataInputDebugFunction(ByVal InputLine As String)
        'Columns are layed out in memory matrix as follows:
        '0,0,0 = Name   0,1,0 = Group #,   racerid,1,0 = ET<15 value for averaging   1,1 = ET<15 number of times found   1,2 = ET<15 result,   2,0 = ET<14 value for averaging   2,1 = ET<14 number times...
        Dim InputLineDestructible As String = InputLine
        Dim RacerName As String = ""
        Dim Value1 As String = ""
        Dim Value2 As String = ""
        Dim Value3 As String = ""
        Dim Value4 As String = ""
        Dim Value5 As String = ""
        Dim Value6 As String = ""
        Dim aa As Integer = 0
        'Acquire Racer Name
        aa = InputLine.IndexOf(",")
        RacerName = InputLine.Remove(0, aa + 1)
        aa = RacerName.IndexOf(",")
        RacerName = RacerName.Remove(aa, RacerName.Length - aa)
        'Acquire Race history number used for averaging
        aa = InputLine.LastIndexOf(",")
        Value2 = InputLine.Remove(0, aa + 1)
        'Acquire Racer ET value for averaging
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value1 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Racer finish number for averaging
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value3 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire resolution values (4 and 5)
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value4 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Value 5
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value5 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Weight (Value 6)
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value6 = InputLineDestructible.Remove(0, aa + 1)
        NewEventMsg("VALUE 1: " & Value1)
        NewEventMsg("VALUE 2: " & Value2)
        NewEventMsg("VALUE 3: " & Value3)
        NewEventMsg("VALUE 4: " & Value4)
        NewEventMsg("VALUE 5: " & Value5)
        NewEventMsg("VALUE 6: " & Value6)
    End Sub

    Private Sub ProcessDataLines(ByVal InputLine As String)
        'Columns are layed out as follows:
        '0,0,0 = Name   0,1,0 = Group #,   racerid,1,0 = ET<15 value for averaging   1,1 = ET<15 number of times found   1,2 = ET<15 result,   2,0 = ET<14 value for averaging   2,1 = ET<14 number times...
        Dim InputLineDestructible As String = InputLine
        Dim RacerName As String = ""
        Dim Value1 As String = ""
        Dim Value2 As String = ""
        Dim Value3 As String = ""
        Dim Value4 As String = ""
        Dim Value5 As String = ""
        Dim Value6 As String = ""
        Dim aa As Integer = 0
        'Acquire Racer Name
        aa = InputLine.IndexOf(",")
        RacerName = InputLine.Remove(0, aa + 1)
        aa = RacerName.IndexOf(",")
        RacerName = RacerName.Remove(aa, RacerName.Length - aa)
        'Acquire Race history number used for averaging
        aa = InputLine.LastIndexOf(",")
        Value2 = InputLine.Remove(0, aa + 1)
        'Acquire Racer ET value for averaging
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value1 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Racer finish number for averaging
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value3 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire resolution values (4 and 5)
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value4 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Value 5
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value5 = InputLineDestructible.Remove(0, aa + 1)
        'Acquire Weight (Value 6)
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        Value6 = InputLineDestructible.Remove(0, aa + 1)
        If LastProcessedName = RacerName Then
            RacerNameTimesFound = RacerNameTimesFound + 1
        Else
            RacerNameTimesFound = 1
            CurrentRacerNumber = CurrentRacerNumber + 1
            Racers(CurrentRacerNumber, 0, 0) = RacerName
            LastProcessedName = RacerName
            NewEventMsg(RacerName)
        End If
        If CInt(Value1) > 29 And CInt(Value1) < 35 Then
            If CInt(Value2) < 15 Then
                Racers(CurrentRacerNumber, 1, 0) = (CSng(Racers(CurrentRacerNumber, 1, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 1, 1) = Racers(CurrentRacerNumber, 1, 1) + 1
                Racers(CurrentRacerNumber, 1, 2) = (CSng(Racers(CurrentRacerNumber, 1, 0)) / CInt(Racers(CurrentRacerNumber, 1, 1)))
            End If
            If CInt(Value2) < 14 Then
                Racers(CurrentRacerNumber, 2, 0) = (CSng(Racers(CurrentRacerNumber, 2, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 2, 1) = Racers(CurrentRacerNumber, 2, 1) + 1
                Racers(CurrentRacerNumber, 2, 2) = (CSng(Racers(CurrentRacerNumber, 2, 0)) / CInt(Racers(CurrentRacerNumber, 2, 1)))
            End If
            If CInt(Value2) < 13 Then
                Racers(CurrentRacerNumber, 3, 0) = (CSng(Racers(CurrentRacerNumber, 3, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 3, 1) = Racers(CurrentRacerNumber, 3, 1) + 1
                Racers(CurrentRacerNumber, 3, 2) = (CSng(Racers(CurrentRacerNumber, 3, 0)) / CInt(Racers(CurrentRacerNumber, 3, 1)))
            End If
            If CInt(Value2) < 12 Then
                Racers(CurrentRacerNumber, 4, 0) = (CSng(Racers(CurrentRacerNumber, 4, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 4, 1) = Racers(CurrentRacerNumber, 4, 1) + 1
                Racers(CurrentRacerNumber, 4, 2) = (CSng(Racers(CurrentRacerNumber, 4, 0)) / CInt(Racers(CurrentRacerNumber, 4, 1)))
            End If
            If CInt(Value2) < 11 Then
                Racers(CurrentRacerNumber, 5, 0) = (CSng(Racers(CurrentRacerNumber, 5, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 5, 1) = Racers(CurrentRacerNumber, 5, 1) + 1
                Racers(CurrentRacerNumber, 5, 2) = (CSng(Racers(CurrentRacerNumber, 5, 0)) / CInt(Racers(CurrentRacerNumber, 5, 1)))
            End If
            If CInt(Value2) < 10 Then
                Racers(CurrentRacerNumber, 6, 0) = (CSng(Racers(CurrentRacerNumber, 6, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 6, 1) = Racers(CurrentRacerNumber, 6, 1) + 1
                Racers(CurrentRacerNumber, 6, 2) = (CSng(Racers(CurrentRacerNumber, 6, 0)) / CInt(Racers(CurrentRacerNumber, 6, 1)))
            End If
            If CInt(Value2) < 9 Then
                Racers(CurrentRacerNumber, 7, 0) = (CSng(Racers(CurrentRacerNumber, 7, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 7, 1) = Racers(CurrentRacerNumber, 7, 1) + 1
                Racers(CurrentRacerNumber, 7, 2) = (CSng(Racers(CurrentRacerNumber, 7, 0)) / CInt(Racers(CurrentRacerNumber, 7, 1)))
            End If
            If CInt(Value2) < 8 Then
                Racers(CurrentRacerNumber, 8, 0) = (CSng(Racers(CurrentRacerNumber, 8, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 8, 1) = Racers(CurrentRacerNumber, 8, 1) + 1
                Racers(CurrentRacerNumber, 8, 2) = (CSng(Racers(CurrentRacerNumber, 8, 0)) / CInt(Racers(CurrentRacerNumber, 8, 1)))
            End If
            If CInt(Value2) < 7 Then
                Racers(CurrentRacerNumber, 9, 0) = (CSng(Racers(CurrentRacerNumber, 9, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 9, 1) = Racers(CurrentRacerNumber, 9, 1) + 1
                Racers(CurrentRacerNumber, 9, 2) = (CSng(Racers(CurrentRacerNumber, 9, 0)) / CInt(Racers(CurrentRacerNumber, 9, 1)))
            End If
            If CInt(Value2) < 6 Then
                Racers(CurrentRacerNumber, 10, 0) = (CSng(Racers(CurrentRacerNumber, 10, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 10, 1) = Racers(CurrentRacerNumber, 10, 1) + 1
                Racers(CurrentRacerNumber, 10, 2) = (CSng(Racers(CurrentRacerNumber, 10, 0)) / CInt(Racers(CurrentRacerNumber, 10, 1)))
            End If
            If CInt(Value2) < 5 Then
                Racers(CurrentRacerNumber, 11, 0) = (CSng(Racers(CurrentRacerNumber, 11, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 11, 1) = Racers(CurrentRacerNumber, 11, 1) + 1
                Racers(CurrentRacerNumber, 11, 2) = (CSng(Racers(CurrentRacerNumber, 11, 0)) / CInt(Racers(CurrentRacerNumber, 11, 1)))
            End If
            If CInt(Value2) < 4 Then
                Racers(CurrentRacerNumber, 12, 0) = (CSng(Racers(CurrentRacerNumber, 12, 0)) + CSng(Value1))
                Racers(CurrentRacerNumber, 12, 1) = Racers(CurrentRacerNumber, 12, 1) + 1
                Racers(CurrentRacerNumber, 12, 2) = (CSng(Racers(CurrentRacerNumber, 12, 0)) / CInt(Racers(CurrentRacerNumber, 12, 1)))
            End If
        End If
        If CInt(Value3) > 0 And CInt(Value3) < 10 Then
            If CInt(Value2) < 15 Then
                Racers(CurrentRacerNumber, 25, 0) = (CSng(Racers(CurrentRacerNumber, 25, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 25, 1) = Racers(CurrentRacerNumber, 25, 1) + 1
                Racers(CurrentRacerNumber, 25, 2) = (CSng(Racers(CurrentRacerNumber, 25, 0)) / CInt(Racers(CurrentRacerNumber, 25, 1)))
            End If
            If CInt(Value2) < 13 Then
                Racers(CurrentRacerNumber, 26, 0) = (CSng(Racers(CurrentRacerNumber, 26, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 26, 1) = Racers(CurrentRacerNumber, 26, 1) + 1
                Racers(CurrentRacerNumber, 26, 2) = (CSng(Racers(CurrentRacerNumber, 26, 0)) / CInt(Racers(CurrentRacerNumber, 26, 1)))
            End If
            If CInt(Value2) < 11 Then
                Racers(CurrentRacerNumber, 27, 0) = (CSng(Racers(CurrentRacerNumber, 27, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 27, 1) = Racers(CurrentRacerNumber, 27, 1) + 1
                Racers(CurrentRacerNumber, 27, 2) = (CSng(Racers(CurrentRacerNumber, 27, 0)) / CInt(Racers(CurrentRacerNumber, 27, 1)))
            End If
            If CInt(Value2) < 9 Then
                Racers(CurrentRacerNumber, 28, 0) = (CSng(Racers(CurrentRacerNumber, 28, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 28, 1) = Racers(CurrentRacerNumber, 28, 1) + 1
                Racers(CurrentRacerNumber, 28, 2) = (CSng(Racers(CurrentRacerNumber, 28, 0)) / CInt(Racers(CurrentRacerNumber, 28, 1)))
            End If
            If CInt(Value2) < 7 Then
                Racers(CurrentRacerNumber, 29, 0) = (CSng(Racers(CurrentRacerNumber, 29, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 29, 1) = Racers(CurrentRacerNumber, 29, 1) + 1
                Racers(CurrentRacerNumber, 29, 2) = (CSng(Racers(CurrentRacerNumber, 29, 0)) / CInt(Racers(CurrentRacerNumber, 29, 1)))
            End If
            If CInt(Value2) < 5 Then
                Racers(CurrentRacerNumber, 30, 0) = (CSng(Racers(CurrentRacerNumber, 30, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 30, 1) = Racers(CurrentRacerNumber, 30, 1) + 1
                Racers(CurrentRacerNumber, 30, 2) = (CSng(Racers(CurrentRacerNumber, 30, 0)) / CInt(Racers(CurrentRacerNumber, 30, 1)))
            End If
            If CInt(Value2) < 3 Then
                Racers(CurrentRacerNumber, 31, 0) = (CSng(Racers(CurrentRacerNumber, 31, 0)) + CSng(Value3))
                Racers(CurrentRacerNumber, 31, 1) = Racers(CurrentRacerNumber, 31, 1) + 1
                Racers(CurrentRacerNumber, 31, 2) = (CSng(Racers(CurrentRacerNumber, 31, 0)) / CInt(Racers(CurrentRacerNumber, 31, 1)))
            End If
        End If
        If CInt(Value4) > 0 And CInt(Value4) < 10 Then
            If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                If CInt(Value2) < 15 Then
                    Racers(CurrentRacerNumber, 32, 0) = (CInt(Racers(CurrentRacerNumber, 32, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 32, 1) = (CInt(Racers(CurrentRacerNumber, 32, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 32, 2) = (CInt(Racers(CurrentRacerNumber, 32, 1)) / CInt(Racers(CurrentRacerNumber, 32, 0))) * 100
                End If
                If CInt(Value2) < 13 Then
                    Racers(CurrentRacerNumber, 33, 0) = (CInt(Racers(CurrentRacerNumber, 33, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 33, 1) = (CInt(Racers(CurrentRacerNumber, 33, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 33, 2) = (CInt(Racers(CurrentRacerNumber, 33, 1)) / CInt(Racers(CurrentRacerNumber, 33, 0))) * 100
                End If
                If CInt(Value2) < 11 Then
                    Racers(CurrentRacerNumber, 34, 0) = (CInt(Racers(CurrentRacerNumber, 34, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 34, 1) = (CInt(Racers(CurrentRacerNumber, 34, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 34, 2) = (CInt(Racers(CurrentRacerNumber, 34, 1)) / CInt(Racers(CurrentRacerNumber, 34, 0))) * 100
                End If
                If CInt(Value2) < 9 Then
                    Racers(CurrentRacerNumber, 35, 0) = (CInt(Racers(CurrentRacerNumber, 35, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 35, 1) = (CInt(Racers(CurrentRacerNumber, 35, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 35, 2) = (CInt(Racers(CurrentRacerNumber, 35, 1)) / CInt(Racers(CurrentRacerNumber, 35, 0))) * 100
                End If
                If CInt(Value2) < 7 Then
                    Racers(CurrentRacerNumber, 36, 0) = (CInt(Racers(CurrentRacerNumber, 36, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 36, 1) = (CInt(Racers(CurrentRacerNumber, 36, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 36, 2) = (CInt(Racers(CurrentRacerNumber, 36, 1)) / CInt(Racers(CurrentRacerNumber, 36, 0))) * 100
                End If
                If CInt(Value2) < 5 Then
                    Racers(CurrentRacerNumber, 37, 0) = (CInt(Racers(CurrentRacerNumber, 37, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 37, 1) = (CInt(Racers(CurrentRacerNumber, 37, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 37, 2) = (CInt(Racers(CurrentRacerNumber, 37, 1)) / CInt(Racers(CurrentRacerNumber, 37, 0))) * 100
                End If
                If CInt(Value2) < 3 Then
                    Racers(CurrentRacerNumber, 38, 0) = (CInt(Racers(CurrentRacerNumber, 38, 0)) + CInt(Value4))
                    Racers(CurrentRacerNumber, 38, 1) = (CInt(Racers(CurrentRacerNumber, 38, 1)) + CInt(Value5))
                    Racers(CurrentRacerNumber, 38, 2) = (CInt(Racers(CurrentRacerNumber, 38, 1)) / CInt(Racers(CurrentRacerNumber, 38, 0))) * 100
                End If
            End If
        End If
        If CInt(Value4) > 0 And CInt(Value4) < 10 Then
            If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                If CInt(Value2) < 15 Then
                    Racers(CurrentRacerNumber, 39, 0) = (CSng(Racers(CurrentRacerNumber, 39, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 39, 1) = Racers(CurrentRacerNumber, 39, 1) + 1
                    Racers(CurrentRacerNumber, 39, 2) = (CSng(Racers(CurrentRacerNumber, 39, 0)) / CInt(Racers(CurrentRacerNumber, 39, 1)))
                End If
                If CInt(Value2) < 13 Then
                    Racers(CurrentRacerNumber, 40, 0) = (CSng(Racers(CurrentRacerNumber, 40, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 40, 1) = Racers(CurrentRacerNumber, 40, 1) + 1
                    Racers(CurrentRacerNumber, 40, 2) = (CSng(Racers(CurrentRacerNumber, 40, 0)) / CInt(Racers(CurrentRacerNumber, 40, 1)))
                End If
                If CInt(Value2) < 11 Then
                    Racers(CurrentRacerNumber, 41, 0) = (CSng(Racers(CurrentRacerNumber, 41, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 41, 1) = Racers(CurrentRacerNumber, 41, 1) + 1
                    Racers(CurrentRacerNumber, 41, 2) = (CSng(Racers(CurrentRacerNumber, 41, 0)) / CInt(Racers(CurrentRacerNumber, 41, 1)))
                End If
                If CInt(Value2) < 9 Then
                    Racers(CurrentRacerNumber, 42, 0) = (CSng(Racers(CurrentRacerNumber, 42, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 42, 1) = Racers(CurrentRacerNumber, 42, 1) + 1
                    Racers(CurrentRacerNumber, 42, 2) = (CSng(Racers(CurrentRacerNumber, 42, 0)) / CInt(Racers(CurrentRacerNumber, 42, 1)))
                End If
                If CInt(Value2) < 7 Then
                    Racers(CurrentRacerNumber, 43, 0) = (CSng(Racers(CurrentRacerNumber, 43, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 43, 1) = Racers(CurrentRacerNumber, 43, 1) + 1
                    Racers(CurrentRacerNumber, 43, 2) = (CSng(Racers(CurrentRacerNumber, 43, 0)) / CInt(Racers(CurrentRacerNumber, 43, 1)))
                End If
                If CInt(Value2) < 5 Then
                    Racers(CurrentRacerNumber, 44, 0) = (CSng(Racers(CurrentRacerNumber, 44, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 44, 1) = Racers(CurrentRacerNumber, 44, 1) + 1
                    Racers(CurrentRacerNumber, 44, 2) = (CSng(Racers(CurrentRacerNumber, 44, 0)) / CInt(Racers(CurrentRacerNumber, 44, 1)))
                End If
                If CInt(Value2) < 3 Then
                    Racers(CurrentRacerNumber, 45, 0) = (CSng(Racers(CurrentRacerNumber, 45, 0)) + CSng(Value4))
                    Racers(CurrentRacerNumber, 45, 1) = Racers(CurrentRacerNumber, 45, 1) + 1
                    Racers(CurrentRacerNumber, 45, 2) = (CSng(Racers(CurrentRacerNumber, 45, 0)) / CInt(Racers(CurrentRacerNumber, 45, 1)))
                End If
            End If
        End If
        If CInt(Value4) > 0 And CInt(Value4) < 10 Then
            If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                If CInt(Value2) < 15 Then
                    Racers(CurrentRacerNumber, 46, 0) = (CSng(Racers(CurrentRacerNumber, 46, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 46, 1) = Racers(CurrentRacerNumber, 46, 1) + 1
                    Racers(CurrentRacerNumber, 46, 2) = (CSng(Racers(CurrentRacerNumber, 46, 0)) / CInt(Racers(CurrentRacerNumber, 46, 1)))
                End If
                If CInt(Value2) < 13 Then
                    Racers(CurrentRacerNumber, 47, 0) = (CSng(Racers(CurrentRacerNumber, 47, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 47, 1) = Racers(CurrentRacerNumber, 47, 1) + 1
                    Racers(CurrentRacerNumber, 47, 2) = (CSng(Racers(CurrentRacerNumber, 47, 0)) / CInt(Racers(CurrentRacerNumber, 47, 1)))
                End If
                If CInt(Value2) < 11 Then
                    Racers(CurrentRacerNumber, 48, 0) = (CSng(Racers(CurrentRacerNumber, 48, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 48, 1) = Racers(CurrentRacerNumber, 48, 1) + 1
                    Racers(CurrentRacerNumber, 48, 2) = (CSng(Racers(CurrentRacerNumber, 48, 0)) / CInt(Racers(CurrentRacerNumber, 48, 1)))
                End If
                If CInt(Value2) < 9 Then
                    Racers(CurrentRacerNumber, 49, 0) = (CSng(Racers(CurrentRacerNumber, 49, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 49, 1) = Racers(CurrentRacerNumber, 49, 1) + 1
                    Racers(CurrentRacerNumber, 49, 2) = (CSng(Racers(CurrentRacerNumber, 49, 0)) / CInt(Racers(CurrentRacerNumber, 49, 1)))
                End If
                If CInt(Value2) < 7 Then
                    Racers(CurrentRacerNumber, 50, 0) = (CSng(Racers(CurrentRacerNumber, 50, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 50, 1) = Racers(CurrentRacerNumber, 50, 1) + 1
                    Racers(CurrentRacerNumber, 50, 2) = (CSng(Racers(CurrentRacerNumber, 50, 0)) / CInt(Racers(CurrentRacerNumber, 50, 1)))
                End If
                If CInt(Value2) < 5 Then
                    Racers(CurrentRacerNumber, 51, 0) = (CSng(Racers(CurrentRacerNumber, 51, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 51, 1) = Racers(CurrentRacerNumber, 51, 1) + 1
                    Racers(CurrentRacerNumber, 51, 2) = (CSng(Racers(CurrentRacerNumber, 51, 0)) / CInt(Racers(CurrentRacerNumber, 51, 1)))
                End If
                If CInt(Value2) < 3 Then
                    Racers(CurrentRacerNumber, 52, 0) = (CSng(Racers(CurrentRacerNumber, 52, 0)) + CSng(Value5))
                    Racers(CurrentRacerNumber, 52, 1) = Racers(CurrentRacerNumber, 52, 1) + 1
                    Racers(CurrentRacerNumber, 52, 2) = (CSng(Racers(CurrentRacerNumber, 52, 0)) / CInt(Racers(CurrentRacerNumber, 52, 1)))
                End If
            End If
        End If
        'Average weight for last 5 races
        If CInt(Value6) > 40 And CInt(Value6) < 105 Then
            If CInt(Value2) < 5 Then
                Racers(CurrentRacerNumber, 53, 0) = (CSng(Racers(CurrentRacerNumber, 53, 0)) + CSng(Value6))
                Racers(CurrentRacerNumber, 53, 1) = Racers(CurrentRacerNumber, 53, 1) + 1
                Racers(CurrentRacerNumber, 53, 2) = (CSng(Racers(CurrentRacerNumber, 53, 0)) / CInt(Racers(CurrentRacerNumber, 53, 1)))
            End If
        End If
        'Break Points
        If CInt(Value5) > 0 And CInt(Value5) < 4 Then
            'Add points for breaking 1-3
            If Value5 = 1 Then Racers(CurrentRacerNumber, 54, 0) = (CSng(Racers(CurrentRacerNumber, 54, 0))) + CSng(20)
            If Value5 = 2 Then Racers(CurrentRacerNumber, 54, 0) = (CSng(Racers(CurrentRacerNumber, 54, 0))) + CSng(10)
            If Value5 = 3 Then Racers(CurrentRacerNumber, 54, 0) = (CSng(Racers(CurrentRacerNumber, 54, 0))) + CSng(5)
            'Conditional Averaging of finish
            Racers(CurrentRacerNumber, 55, 0) = (CSng(Racers(CurrentRacerNumber, 55, 0)) + CSng(Value3))
            Racers(CurrentRacerNumber, 55, 1) = Racers(CurrentRacerNumber, 55, 1) + 1
            Racers(CurrentRacerNumber, 55, 2) = (CSng(Racers(CurrentRacerNumber, 55, 0)) / CInt(Racers(CurrentRacerNumber, 55, 1)))
        End If
    End Sub

    Private Sub ReadGroupingFile(ByVal InputFile As String)
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied group file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs3 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr3 As New StreamReader(fs3)
            If _LastOffset3 < sr3.BaseStream.Length Then
                sr3.BaseStream.Seek(_LastOffset3, SeekOrigin.Begin)
                While sr3.Peek() <> -1
                    Dim read As String = sr3.ReadLine()
                    If read.Length > 0 Then
                        If CurrentGroupMembers = 8 Then
                            CurrentGroupMembers = 1
                            CurrentGroup = CurrentGroup + 1
                        Else
                            CurrentGroupMembers = CurrentGroupMembers + 1
                        End If
                        AssignGroup(read, CurrentGroup)
                    End If
                End While
                _LastOffset3 = sr3.BaseStream.Position
            End If
            sr3.Close()
            fs3.Close()
            TotalGroups = CurrentGroup
            NewEventMsg("Finished reading grouping file, assigned " & CStr(TotalGroups) & " groups")
            CurrentGroup = 0
            CurrentGroupMembers = 0
        End If
    End Sub

    Private Sub AssignGroup(ByVal InputLine As String, ByVal GroupNum As Integer)
        Dim RacerName As String = InputLine
        For o = 0 To UniqueRacerNamesFound
            If RacerName = Racers(o, 0, 0) Then
                'If name is equal to input line of group list, then add that name to the current group
                Racers(o, 0, 1) = CurrentGroup
            End If
        Next o
    End Sub

    Private Sub CalculateRanksInGroup()
        NewEventMsg("Calculating ranks and comparative averages...")
        CurrentGroupMembers = 0
        For p = 1 To TotalGroups 'For Each group from 0 to the total amount of groups created
            Dim GroupMembers(MaxGroupSize) As Integer
            For q = 0 To UniqueRacerNamesFound 'For each racer in memory, get a list of the members of the current group as an array in GroupMembers(1 thru 8)
                If Racers(q, 0, 1) = p Then
                    If CurrentGroupMembers = 8 Then
                        CurrentGroupMembers = 1
                    Else
                        CurrentGroupMembers = CurrentGroupMembers + 1
                    End If
                    GroupMembers(CurrentGroupMembers) = q
                End If
            Next q
            'Once every member of the group is obtained
            For t = 0 To 11 'For each ET<?, rank each racer in each tier
                For h = 1 To GroupMembers.Count - 1 'For each member of the group
                    Dim CurrentRank As Integer = 1
                    For i = 1 To GroupMembers.Count - 1 'For every other member of the group
                        If Racers(GroupMembers(h), (1 + t), 2) > Racers(GroupMembers(i), (1 + t), 2) Then
                            CurrentRank = CurrentRank + 1
                        End If
                    Next i
                    Racers(GroupMembers(h), (13 + t), 2) = CurrentRank
                Next h
            Next t
        Next p
        NewEventMsg("Finished!")
    End Sub

    Private Sub UpdateDataGridView(ByVal InputFile As String)
        Dim LinesReadIn As Integer = 0
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs4 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr4 As New StreamReader(fs4)
            If _LastOffset4 < sr4.BaseStream.Length Then
                sr4.BaseStream.Seek(_LastOffset4, SeekOrigin.Begin)
                While sr4.Peek() <> -1
                    Dim read As String = sr4.ReadLine()
                    If read.Length > 0 Then
                        LinesReadIn = LinesReadIn + 1
                        AddLineToGrid(read)
                    End If
                End While
                _LastOffset4 = sr4.BaseStream.Position
            End If
            sr4.Close()
            fs4.Close()
            NewEventMsg("VIEW: Used " & LinesReadIn & " lines to populate Grid View from file: " & ViewFile)
        End If
    End Sub

    Private Sub AddLineToGrid(ByVal InputLine As String)
        Dim InputLineDestructible As String = InputLine
        Dim NewValues(DelimiterCount) As String
        Dim newrow As String()
        Dim ac As Integer = 0
        For x = 0 To DelimiterCount
            If InputLineDestructible.IndexOf(",") > -1 Then
                ac = InputLineDestructible.IndexOf(",")
                NewValues(x) = InputLineDestructible.Remove(ac, InputLineDestructible.Length - ac)
                InputLineDestructible = InputLineDestructible.Remove(0, ac + 1)
            Else
                If InputLineDestructible.Length > 0 Then
                    NewValues(x) = InputLineDestructible
                    InputLineDestructible = ""
                End If
            End If
        Next x
        newrow = New String() {NewValues(0), NewValues(1), NewValues(2), NewValues(3), NewValues(4), NewValues(5), NewValues(6), NewValues(7), NewValues(8), NewValues(9), NewValues(10), NewValues(11), NewValues(12), NewValues(13), NewValues(14), NewValues(15), NewValues(16), NewValues(17), NewValues(18), NewValues(19), NewValues(20), NewValues(21), NewValues(22), NewValues(23), NewValues(24), NewValues(25), NewValues(26), NewValues(27), NewValues(28), NewValues(29), NewValues(30), NewValues(31), NewValues(32), NewValues(33), NewValues(34), NewValues(35), NewValues(36), NewValues(37), NewValues(38), NewValues(39), NewValues(40), NewValues(41), NewValues(42), NewValues(43), NewValues(44), NewValues(45), NewValues(46), NewValues(47), NewValues(48), NewValues(49), NewValues(50), NewValues(51), NewValues(52), NewValues(53), NewValues(54), NewValues(55), NewValues(56), NewValues(57)}
        DataGridView1.Rows.Add(newrow)
    End Sub


    '########## Back-Testing Functions ##########

    Private Sub CountLinesInFile(ByVal InputFile As String)
        LinesInFile = 0
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs5 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr5 As New StreamReader(fs5)
            If _LastOffset5 < sr5.BaseStream.Length Then
                sr5.BaseStream.Seek(_LastOffset5, SeekOrigin.Begin)
                While sr5.Peek() <> -1
                    Dim read As String = sr5.ReadLine()
                    If read.Length > 0 Then
                        LinesInFile = LinesInFile + 1
                        CountRacers(read)
                    End If
                End While
                _LastOffset = sr5.BaseStream.Position
            End If
            sr5.Close()
            fs5.Close()
            NewEventMsg("COUNT: " & UniqueRacerNamesFound & " lines found in file.")
        End If
    End Sub

    Private Sub ImportBackTestData(ByVal InputFile As String)
        CurrentLineNumber = 0
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs6 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr6 As New StreamReader(fs6)
            If _LastOffset6 < sr6.BaseStream.Length Then
                sr6.BaseStream.Seek(_LastOffset6, SeekOrigin.Begin)
                While sr6.Peek() <> -1
                    Dim read As String = sr6.ReadLine()
                    If read.Length <> 0 Then
                        BackTestReadData(read)
                    End If
                    CurrentLineNumber = CurrentLineNumber + 1
                End While
                _LastOffset2 = sr6.BaseStream.Position
            End If
            sr6.Close()
            fs6.Close()
            _LastOffset2 = 0
            NewEventMsg("FINISHED: Read " & CurrentLineNumber & " lines, and parsed them into memory.")
        End If
    End Sub

    Private Sub BackTestReadData(ByVal InputLine As String)
        'Reads each line of input file into memory array
        Dim InputLineDestructible As String = InputLine
        Dim aa As Integer = 0
        'x,0,0 = date, x,1,0 = name, x,2,0 = datacolumn ...
        aa = InputLineDestructible.IndexOf(",")
        RaceHistory(CurrentLineNumber, 0) = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        For ab = 1 To 13
            aa = InputLineDestructible.IndexOf(",")
            RaceHistory(CurrentLineNumber, ab) = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
            InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        Next ab
        RaceHistory(CurrentLineNumber, 14) = InputLineDestructible
    End Sub

    Private Sub ReImportBackTestOutput(ByVal InputFile As String)
        CurrentLineNumber = 0
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs6 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr6 As New StreamReader(fs6)
            If _LastOffset6 < sr6.BaseStream.Length Then
                sr6.BaseStream.Seek(_LastOffset6, SeekOrigin.Begin)
                While sr6.Peek() <> -1
                    Dim read As String = sr6.ReadLine()
                    If read.Length <> 0 Then
                        BackTestReadData2(read)
                    End If
                    CurrentLineNumber = CurrentLineNumber + 1
                End While
                _LastOffset2 = sr6.BaseStream.Position
            End If
            sr6.Close()
            fs6.Close()
            _LastOffset2 = 0
            NewEventMsg("FINISHED: Read " & CurrentLineNumber & " lines, and parsed them into memory.")
        End If
    End Sub

    Private Sub BackTestReadData2(ByVal InputLine As String)
        'Reads each line of input file into memory array
        Dim InputLineDestructible As String = InputLine
        Dim aa As Integer = 0
        'x,0 = date, x,1 = name, x,2 = datacolumn ...
        aa = InputLineDestructible.IndexOf(",")
        OutputFileArray(CurrentLineNumber, 0) = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        For ab = 1 To 54
            aa = InputLineDestructible.IndexOf(",")
            OutputFileArray(CurrentLineNumber, ab) = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
            InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        Next ab
        OutputFileArray(CurrentLineNumber, 55) = InputLineDestructible
    End Sub

    Private Sub ComputeBackTestData()
        For CurrLine = 1 To LinesInFile 'For each line in the input file
            If LastProcessedName = RaceHistory(CurrLine, 1) Then
                RacerNameTimesFound = RacerNameTimesFound + 1
            Else
                RacerNameTimesFound = 1
                CurrentRacerNumber = CurrentRacerNumber + 1
                Racers(CurrentRacerNumber - 1, 0, 0) = RaceHistory(CurrLine, 1)
                LastProcessedName = RaceHistory(CurrLine, 1)
                NewEventMsg("Computing history for: " & CStr(RaceHistory(CurrLine, 1)) & " starting: " & CStr(RaceHistory(CurrLine, 0)))
            End If
            If RacerNameTimesFound > 4 Then
                'Empty previous results before proceeding
                For bb = 1 To DelimiterCount
                    For bc = 0 To 2
                        Racers(CurrentRacerNumber, bb, bc) = 00
                    Next bc
                Next bb
                For ac = 1 To RacerNameTimesFound
                    'Adapt our variables and memory format to the same variable scheme because copy paste = life
                    Dim Value1 As String = RaceHistory(CurrLine - ac, 14)
                    Dim Value2 As String = ac
                    Dim Value3 As String = RaceHistory(CurrLine - ac, 12)
                    Dim Value4 As String = RaceHistory(CurrLine - ac, 10)
                    Dim Value5 As String = RaceHistory(CurrLine - ac, 9)

                    If CInt(Value1) > 29 And CInt(Value1) < 35 Then
                        If CInt(Value2) > 15 Then
                            Racers(CurrentRacerNumber, 1, 0) = (CSng(Racers(CurrentRacerNumber, 1, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 1, 1) = Racers(CurrentRacerNumber, 1, 1) + 1
                            Racers(CurrentRacerNumber, 1, 2) = (CSng(Racers(CurrentRacerNumber, 1, 0)) / CInt(Racers(CurrentRacerNumber, 1, 1)))
                        End If
                        If CInt(Value2) > 14 Then
                            Racers(CurrentRacerNumber, 2, 0) = (CSng(Racers(CurrentRacerNumber, 2, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 2, 1) = Racers(CurrentRacerNumber, 2, 1) + 1
                            Racers(CurrentRacerNumber, 2, 2) = (CSng(Racers(CurrentRacerNumber, 2, 0)) / CInt(Racers(CurrentRacerNumber, 2, 1)))
                        End If
                        If CInt(Value2) > 13 Then
                            Racers(CurrentRacerNumber, 3, 0) = (CSng(Racers(CurrentRacerNumber, 3, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 3, 1) = Racers(CurrentRacerNumber, 3, 1) + 1
                            Racers(CurrentRacerNumber, 3, 2) = (CSng(Racers(CurrentRacerNumber, 3, 0)) / CInt(Racers(CurrentRacerNumber, 3, 1)))
                        End If
                        If CInt(Value2) > 12 Then
                            Racers(CurrentRacerNumber, 4, 0) = (CSng(Racers(CurrentRacerNumber, 4, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 4, 1) = Racers(CurrentRacerNumber, 4, 1) + 1
                            Racers(CurrentRacerNumber, 4, 2) = (CSng(Racers(CurrentRacerNumber, 4, 0)) / CInt(Racers(CurrentRacerNumber, 4, 1)))
                        End If
                        If CInt(Value2) > 11 Then
                            Racers(CurrentRacerNumber, 5, 0) = (CSng(Racers(CurrentRacerNumber, 5, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 5, 1) = Racers(CurrentRacerNumber, 5, 1) + 1
                            Racers(CurrentRacerNumber, 5, 2) = (CSng(Racers(CurrentRacerNumber, 5, 0)) / CInt(Racers(CurrentRacerNumber, 5, 1)))
                        End If
                        If CInt(Value2) > 10 Then
                            Racers(CurrentRacerNumber, 6, 0) = (CSng(Racers(CurrentRacerNumber, 6, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 6, 1) = Racers(CurrentRacerNumber, 6, 1) + 1
                            Racers(CurrentRacerNumber, 6, 2) = (CSng(Racers(CurrentRacerNumber, 6, 0)) / CInt(Racers(CurrentRacerNumber, 6, 1)))
                        End If
                        If CInt(Value2) > 9 Then
                            Racers(CurrentRacerNumber, 7, 0) = (CSng(Racers(CurrentRacerNumber, 7, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 7, 1) = Racers(CurrentRacerNumber, 7, 1) + 1
                            Racers(CurrentRacerNumber, 7, 2) = (CSng(Racers(CurrentRacerNumber, 7, 0)) / CInt(Racers(CurrentRacerNumber, 7, 1)))
                        End If
                        If CInt(Value2) > 8 Then
                            Racers(CurrentRacerNumber, 8, 0) = (CSng(Racers(CurrentRacerNumber, 8, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 8, 1) = Racers(CurrentRacerNumber, 8, 1) + 1
                            Racers(CurrentRacerNumber, 8, 2) = (CSng(Racers(CurrentRacerNumber, 8, 0)) / CInt(Racers(CurrentRacerNumber, 8, 1)))
                        End If
                        If CInt(Value2) > 7 Then
                            Racers(CurrentRacerNumber, 9, 0) = (CSng(Racers(CurrentRacerNumber, 9, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 9, 1) = Racers(CurrentRacerNumber, 9, 1) + 1
                            Racers(CurrentRacerNumber, 9, 2) = (CSng(Racers(CurrentRacerNumber, 9, 0)) / CInt(Racers(CurrentRacerNumber, 9, 1)))
                        End If
                        If CInt(Value2) > 6 Then
                            Racers(CurrentRacerNumber, 10, 0) = (CSng(Racers(CurrentRacerNumber, 10, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 10, 1) = Racers(CurrentRacerNumber, 10, 1) + 1
                            Racers(CurrentRacerNumber, 10, 2) = (CSng(Racers(CurrentRacerNumber, 10, 0)) / CInt(Racers(CurrentRacerNumber, 10, 1)))
                        End If
                        If CInt(Value2) > 5 Then
                            Racers(CurrentRacerNumber, 11, 0) = (CSng(Racers(CurrentRacerNumber, 11, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 11, 1) = Racers(CurrentRacerNumber, 11, 1) + 1
                            Racers(CurrentRacerNumber, 11, 2) = (CSng(Racers(CurrentRacerNumber, 11, 0)) / CInt(Racers(CurrentRacerNumber, 11, 1)))
                        End If
                        If CInt(Value2) > 4 Then
                            Racers(CurrentRacerNumber, 12, 0) = (CSng(Racers(CurrentRacerNumber, 12, 0)) + CSng(Value1))
                            Racers(CurrentRacerNumber, 12, 1) = Racers(CurrentRacerNumber, 12, 1) + 1
                            Racers(CurrentRacerNumber, 12, 2) = (CSng(Racers(CurrentRacerNumber, 12, 0)) / CInt(Racers(CurrentRacerNumber, 12, 1)))
                        End If
                    End If
                    If CInt(Value3) > 0 And CInt(Value3) < 10 Then
                        If CInt(Value2) < 15 Then
                            Racers(CurrentRacerNumber, 25, 0) = (CSng(Racers(CurrentRacerNumber, 25, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 25, 1) = Racers(CurrentRacerNumber, 25, 1) + 1
                            Racers(CurrentRacerNumber, 25, 2) = (CSng(Racers(CurrentRacerNumber, 25, 0)) / CInt(Racers(CurrentRacerNumber, 25, 1)))
                        End If
                        If CInt(Value2) < 13 Then
                            Racers(CurrentRacerNumber, 26, 0) = (CSng(Racers(CurrentRacerNumber, 26, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 26, 1) = Racers(CurrentRacerNumber, 26, 1) + 1
                            Racers(CurrentRacerNumber, 26, 2) = (CSng(Racers(CurrentRacerNumber, 26, 0)) / CInt(Racers(CurrentRacerNumber, 26, 1)))
                        End If
                        If CInt(Value2) < 11 Then
                            Racers(CurrentRacerNumber, 27, 0) = (CSng(Racers(CurrentRacerNumber, 27, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 27, 1) = Racers(CurrentRacerNumber, 27, 1) + 1
                            Racers(CurrentRacerNumber, 27, 2) = (CSng(Racers(CurrentRacerNumber, 27, 0)) / CInt(Racers(CurrentRacerNumber, 27, 1)))
                        End If
                        If CInt(Value2) < 9 Then
                            Racers(CurrentRacerNumber, 28, 0) = (CSng(Racers(CurrentRacerNumber, 28, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 28, 1) = Racers(CurrentRacerNumber, 28, 1) + 1
                            Racers(CurrentRacerNumber, 28, 2) = (CSng(Racers(CurrentRacerNumber, 28, 0)) / CInt(Racers(CurrentRacerNumber, 28, 1)))
                        End If
                        If CInt(Value2) < 7 Then
                            Racers(CurrentRacerNumber, 29, 0) = (CSng(Racers(CurrentRacerNumber, 29, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 29, 1) = Racers(CurrentRacerNumber, 29, 1) + 1
                            Racers(CurrentRacerNumber, 29, 2) = (CSng(Racers(CurrentRacerNumber, 29, 0)) / CInt(Racers(CurrentRacerNumber, 29, 1)))
                        End If
                        If CInt(Value2) < 5 Then
                            Racers(CurrentRacerNumber, 30, 0) = (CSng(Racers(CurrentRacerNumber, 30, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 30, 1) = Racers(CurrentRacerNumber, 30, 1) + 1
                            Racers(CurrentRacerNumber, 30, 2) = (CSng(Racers(CurrentRacerNumber, 30, 0)) / CInt(Racers(CurrentRacerNumber, 30, 1)))
                        End If
                        If CInt(Value2) < 3 Then
                            Racers(CurrentRacerNumber, 31, 0) = (CSng(Racers(CurrentRacerNumber, 31, 0)) + CSng(Value3))
                            Racers(CurrentRacerNumber, 31, 1) = Racers(CurrentRacerNumber, 31, 1) + 1
                            Racers(CurrentRacerNumber, 31, 2) = (CSng(Racers(CurrentRacerNumber, 31, 0)) / CInt(Racers(CurrentRacerNumber, 31, 1)))
                        End If
                    End If
                    If CInt(Value4) > 0 And CInt(Value4) < 10 Then
                        If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                            If CInt(Value2) < 15 Then
                                Racers(CurrentRacerNumber, 32, 0) = (CInt(Racers(CurrentRacerNumber, 32, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 32, 1) = (CInt(Racers(CurrentRacerNumber, 32, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 32, 2) = (CInt(Racers(CurrentRacerNumber, 32, 1)) / CInt(Racers(CurrentRacerNumber, 32, 0))) * 100
                            End If
                            If CInt(Value2) < 13 Then
                                Racers(CurrentRacerNumber, 33, 0) = (CInt(Racers(CurrentRacerNumber, 33, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 33, 1) = (CInt(Racers(CurrentRacerNumber, 33, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 33, 2) = (CInt(Racers(CurrentRacerNumber, 33, 1)) / CInt(Racers(CurrentRacerNumber, 33, 0))) * 100
                            End If
                            If CInt(Value2) < 11 Then
                                Racers(CurrentRacerNumber, 34, 0) = (CInt(Racers(CurrentRacerNumber, 34, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 34, 1) = (CInt(Racers(CurrentRacerNumber, 34, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 34, 2) = (CInt(Racers(CurrentRacerNumber, 34, 1)) / CInt(Racers(CurrentRacerNumber, 34, 0))) * 100
                            End If
                            If CInt(Value2) < 9 Then
                                Racers(CurrentRacerNumber, 35, 0) = (CInt(Racers(CurrentRacerNumber, 35, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 35, 1) = (CInt(Racers(CurrentRacerNumber, 35, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 35, 2) = (CInt(Racers(CurrentRacerNumber, 35, 1)) / CInt(Racers(CurrentRacerNumber, 35, 0))) * 100
                            End If
                            If CInt(Value2) < 7 Then
                                Racers(CurrentRacerNumber, 36, 0) = (CInt(Racers(CurrentRacerNumber, 36, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 36, 1) = (CInt(Racers(CurrentRacerNumber, 36, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 36, 2) = (CInt(Racers(CurrentRacerNumber, 36, 1)) / CInt(Racers(CurrentRacerNumber, 36, 0))) * 100
                            End If
                            If CInt(Value2) < 5 Then
                                Racers(CurrentRacerNumber, 37, 0) = (CInt(Racers(CurrentRacerNumber, 37, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 37, 1) = (CInt(Racers(CurrentRacerNumber, 37, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 37, 2) = (CInt(Racers(CurrentRacerNumber, 37, 1)) / CInt(Racers(CurrentRacerNumber, 37, 0))) * 100
                            End If
                            If CInt(Value2) < 3 Then
                                Racers(CurrentRacerNumber, 38, 0) = (CInt(Racers(CurrentRacerNumber, 38, 0)) + CInt(Value4))
                                Racers(CurrentRacerNumber, 38, 1) = (CInt(Racers(CurrentRacerNumber, 38, 1)) + CInt(Value5))
                                Racers(CurrentRacerNumber, 38, 2) = (CInt(Racers(CurrentRacerNumber, 38, 1)) / CInt(Racers(CurrentRacerNumber, 38, 0))) * 100
                            End If
                        End If
                    End If
                    If CInt(Value4) > 0 And CInt(Value4) < 10 Then
                        If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                            If CInt(Value2) < 15 Then
                                Racers(CurrentRacerNumber, 39, 0) = (CSng(Racers(CurrentRacerNumber, 39, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 39, 1) = Racers(CurrentRacerNumber, 39, 1) + 1
                                Racers(CurrentRacerNumber, 39, 2) = (CSng(Racers(CurrentRacerNumber, 39, 0)) / CInt(Racers(CurrentRacerNumber, 39, 1)))
                            End If
                            If CInt(Value2) < 13 Then
                                Racers(CurrentRacerNumber, 40, 0) = (CSng(Racers(CurrentRacerNumber, 40, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 40, 1) = Racers(CurrentRacerNumber, 40, 1) + 1
                                Racers(CurrentRacerNumber, 40, 2) = (CSng(Racers(CurrentRacerNumber, 40, 0)) / CInt(Racers(CurrentRacerNumber, 40, 1)))
                            End If
                            If CInt(Value2) < 11 Then
                                Racers(CurrentRacerNumber, 41, 0) = (CSng(Racers(CurrentRacerNumber, 41, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 41, 1) = Racers(CurrentRacerNumber, 41, 1) + 1
                                Racers(CurrentRacerNumber, 41, 2) = (CSng(Racers(CurrentRacerNumber, 41, 0)) / CInt(Racers(CurrentRacerNumber, 41, 1)))
                            End If
                            If CInt(Value2) < 9 Then
                                Racers(CurrentRacerNumber, 42, 0) = (CSng(Racers(CurrentRacerNumber, 42, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 42, 1) = Racers(CurrentRacerNumber, 42, 1) + 1
                                Racers(CurrentRacerNumber, 42, 2) = (CSng(Racers(CurrentRacerNumber, 42, 0)) / CInt(Racers(CurrentRacerNumber, 42, 1)))
                            End If
                            If CInt(Value2) < 7 Then
                                Racers(CurrentRacerNumber, 43, 0) = (CSng(Racers(CurrentRacerNumber, 43, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 43, 1) = Racers(CurrentRacerNumber, 43, 1) + 1
                                Racers(CurrentRacerNumber, 43, 2) = (CSng(Racers(CurrentRacerNumber, 43, 0)) / CInt(Racers(CurrentRacerNumber, 43, 1)))
                            End If
                            If CInt(Value2) < 5 Then
                                Racers(CurrentRacerNumber, 44, 0) = (CSng(Racers(CurrentRacerNumber, 44, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 44, 1) = Racers(CurrentRacerNumber, 44, 1) + 1
                                Racers(CurrentRacerNumber, 44, 2) = (CSng(Racers(CurrentRacerNumber, 44, 0)) / CInt(Racers(CurrentRacerNumber, 44, 1)))
                            End If
                            If CInt(Value2) < 3 Then
                                Racers(CurrentRacerNumber, 45, 0) = (CSng(Racers(CurrentRacerNumber, 45, 0)) + CSng(Value4))
                                Racers(CurrentRacerNumber, 45, 1) = Racers(CurrentRacerNumber, 45, 1) + 1
                                Racers(CurrentRacerNumber, 45, 2) = (CSng(Racers(CurrentRacerNumber, 45, 0)) / CInt(Racers(CurrentRacerNumber, 45, 1)))
                            End If
                        End If
                    End If
                    If CInt(Value4) > 0 And CInt(Value4) < 10 Then
                        If CInt(Value5) > 0 And CInt(Value5) < 10 Then
                            If CInt(Value2) < 15 Then
                                Racers(CurrentRacerNumber, 46, 0) = (CSng(Racers(CurrentRacerNumber, 46, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 46, 1) = Racers(CurrentRacerNumber, 46, 1) + 1
                                Racers(CurrentRacerNumber, 46, 2) = (CSng(Racers(CurrentRacerNumber, 46, 0)) / CInt(Racers(CurrentRacerNumber, 46, 1)))
                            End If
                            If CInt(Value2) < 13 Then
                                Racers(CurrentRacerNumber, 47, 0) = (CSng(Racers(CurrentRacerNumber, 47, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 47, 1) = Racers(CurrentRacerNumber, 47, 1) + 1
                                Racers(CurrentRacerNumber, 47, 2) = (CSng(Racers(CurrentRacerNumber, 47, 0)) / CInt(Racers(CurrentRacerNumber, 47, 1)))
                            End If
                            If CInt(Value2) < 11 Then
                                Racers(CurrentRacerNumber, 48, 0) = (CSng(Racers(CurrentRacerNumber, 48, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 48, 1) = Racers(CurrentRacerNumber, 48, 1) + 1
                                Racers(CurrentRacerNumber, 48, 2) = (CSng(Racers(CurrentRacerNumber, 48, 0)) / CInt(Racers(CurrentRacerNumber, 48, 1)))
                            End If
                            If CInt(Value2) < 9 Then
                                Racers(CurrentRacerNumber, 49, 0) = (CSng(Racers(CurrentRacerNumber, 49, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 49, 1) = Racers(CurrentRacerNumber, 49, 1) + 1
                                Racers(CurrentRacerNumber, 49, 2) = (CSng(Racers(CurrentRacerNumber, 49, 0)) / CInt(Racers(CurrentRacerNumber, 49, 1)))
                            End If
                            If CInt(Value2) < 7 Then
                                Racers(CurrentRacerNumber, 50, 0) = (CSng(Racers(CurrentRacerNumber, 50, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 50, 1) = Racers(CurrentRacerNumber, 50, 1) + 1
                                Racers(CurrentRacerNumber, 50, 2) = (CSng(Racers(CurrentRacerNumber, 50, 0)) / CInt(Racers(CurrentRacerNumber, 50, 1)))
                            End If
                            If CInt(Value2) < 5 Then
                                Racers(CurrentRacerNumber, 51, 0) = (CSng(Racers(CurrentRacerNumber, 51, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 51, 1) = Racers(CurrentRacerNumber, 51, 1) + 1
                                Racers(CurrentRacerNumber, 51, 2) = (CSng(Racers(CurrentRacerNumber, 51, 0)) / CInt(Racers(CurrentRacerNumber, 51, 1)))
                            End If
                            If CInt(Value2) < 3 Then
                                Racers(CurrentRacerNumber, 52, 0) = (CSng(Racers(CurrentRacerNumber, 52, 0)) + CSng(Value5))
                                Racers(CurrentRacerNumber, 52, 1) = Racers(CurrentRacerNumber, 52, 1) + 1
                                Racers(CurrentRacerNumber, 52, 2) = (CSng(Racers(CurrentRacerNumber, 52, 0)) / CInt(Racers(CurrentRacerNumber, 52, 1)))
                            End If
                        End If
                    End If
                Next ac
                'Immediately write results out to disk so the same workspace memory can be re-used next cycle (do not store values in memory)
                Using sw As StreamWriter = File.AppendText(SaveFile)
                    sw.WriteLine(CStr(RaceHistory(CurrLine, 1)) & "," & CStr(RaceHistory(CurrLine, 0)) & "," & CStr(RaceHistory(CurrLine, 2)) & "," & CStr(RaceHistory(CurrLine, 3)) & "," & CStr(CDbl(Racers(CurrentRacerNumber, 0, 1))) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 1, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 2, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 3, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 4, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 5, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 6, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 7, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 8, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 9, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 10, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 11, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 12, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 13, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 14, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 15, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 16, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 17, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 18, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 19, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 20, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 21, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 22, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 23, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 24, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 25, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 26, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 27, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 28, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 29, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 30, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 31, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 32, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 33, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 34, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 35, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 36, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 37, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 38, 2)), 0)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 39, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 40, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 41, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 42, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 43, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 44, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 45, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 46, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 47, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 48, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 49, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 50, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 51, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 52, 2)), 2)) & "," & CStr(Math.Round(CDbl(Racers(CurrentRacerNumber, 1, 1)), 2)))
                End Using
            End If
        Next CurrLine
    End Sub

    Private Sub BackTestFindGroupMembers(ByVal SuperLine As Integer)
        Dim LineRacerName As String = RaceHistory(SuperLine, 1)
        Dim LineDate As String = RaceHistory(SuperLine, 0)
        Dim LineRaceTime As String = RaceHistory(SuperLine, 2)
        Dim LineTrackName As String = RaceHistory(SuperLine, 3)
        Dim RaceMembers(MaxGroupSize) As Integer
        CurrentGroupMembers = 0
        Dim ae As Integer = 0
        For ad = 1 To LinesInFile
            If RaceHistory(ad, 0) = LineDate Then
                If RaceHistory(ad, 2) = LineRaceTime Then
                    If RaceHistory(ad, 3) = LineTrackName Then
                        If RaceHistory(ad, 1) = LineRacerName Then
                            'Do Nothing - This is the line we are currently on, or a duplicate of it
                            If ad = SuperLine Then
                                If ae = 8 Then
                                    ae = 1
                                Else
                                    ae = CurrentGroupMembers + 1
                                End If
                                RaceMembers(ae) = ad
                            End If
                        Else
                            If ae = 8 Then
                                ae = 1
                            Else
                                ae = CurrentGroupMembers + 1
                            End If
                            RaceMembers(ae) = ad
                        End If
                    End If
                End If
            End If
        Next ad
        'Once we have all members of the race that occurs on that line, perform our group calculations
        For t = 0 To 11 'For each ET<?, rank each dog in each tier
            For h = 1 To RaceMembers.Count - 1 'For each member of the group
                Dim CurrentRank As Integer = 1
                For i = 1 To RaceMembers.Count - 1 'For every other member of the group
                    If Racers(RaceMembers(h), (6 + t), 2) > Racers(RaceMembers(i), (6 + t), 2) Then
                        CurrentRank = CurrentRank + 1
                    End If
                Next i
                Racers(RaceMembers(h), (17 + t), 2) = CurrentRank
            Next h
        Next t
    End Sub

    '########## Generation Array Functions ##########
    Private Sub CountUniqueRacerNames(ByVal InputFile As String)
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs7 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr7 As New StreamReader(fs7)
            If _LastOffset7 < sr7.BaseStream.Length Then
                sr7.BaseStream.Seek(_LastOffset7, SeekOrigin.Begin)
                While sr7.Peek() <> -1
                    Dim read As String = sr7.ReadLine()
                    If read.Length > 0 Then
                        Dim RacerName As String = ""
                        Dim aa As Integer = 0
                        aa = read.IndexOf(",")
                        RacerName = read.Remove(0, aa + 1)
                        aa = RacerName.IndexOf(",")
                        RacerName = RacerName.Remove(aa, RacerName.Length - aa)
                        Dim isDuplicate As Boolean = False
                        For ac = 0 To UniqueRacerNamesFound
                            If RacerList(ac) = RacerName Then
                                isDuplicate = True
                            End If
                        Next ac
                        If isDuplicate = False Then
                            RacerList(UniqueRacerNamesFound + 1) = RacerName
                            UniqueRacerNamesFound = UniqueRacerNamesFound + 1
                        End If
                        'check if we're running out of list space, if so, add more
                        If UniqueRacerNamesFound >= RacerList.Count - 1 Then
                            RaceDataArraySize = UniqueRacerNamesFound + 250
                            ReDim Preserve RacerList(RaceDataArraySize)
                        End If
                        Application.DoEvents()
                    End If
                End While
                _LastOffset7 = sr7.BaseStream.Position
            End If
            sr7.Close()
            fs7.Close()
            NewEventMsg("COUNT: " & UniqueRacerNamesFound & " unique racers counted.")
            ReDim RacerMatrix(UniqueRacerNamesFound, UniqueRacerNamesFound, 8, 2)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
        ViewFile = OpenFileDialog1.FileName()
        ViewFileTextBox.Text = ViewFile
        NewEventMsg("VIEW: Counting racers in file: " & ViewFile)
        CountUniqueRacerNames(ViewFile)
        ReadDataLines2(ViewFile)
    End Sub

    Private Sub PopulateRaceDataMatrix(ByVal InputLine As String)
        Dim InputLineDestructible As String = InputLine
        Dim RaceDate As String = ""
        Dim RacerName As String = ""
        Dim RaceNumber As String = ""
        Dim RaceTrack As String = ""
        Dim RaceLength As String = ""
        Dim val1 As String = ""
        Dim val2 As String = ""
        Dim val3 As String = ""
        Dim RaceStartBox As String = ""
        Dim RaceBreak As String = ""
        Dim RacerFinish As String = ""
        Dim RacerScore As String = ""
        Dim aa As Integer = 0

        'Acquire score value
        aa = InputLineDestructible.LastIndexOf(",")
        RacerScore = InputLineDestructible.Remove(0, aa + 1)

        'Acquire Finish
        InputLineDestructible = InputLine
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        InputLineDestructible = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        aa = InputLineDestructible.LastIndexOf(",")
        RacerFinish = InputLineDestructible.Remove(0, aa + 1)
        InputLineDestructible = InputLine
        'Race Date
        aa = InputLineDestructible.IndexOf(",")
        RaceDate = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RacerName = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RaceNumber = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RaceTrack = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RaceLength = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        val1 = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        val2 = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        val3 = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RaceStartBox = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)
        InputLineDestructible = InputLineDestructible.Remove(0, aa + 1)
        aa = InputLineDestructible.IndexOf(",")
        RaceBreak = InputLineDestructible.Remove(aa, InputLineDestructible.Length - aa)

        'create an array (10x8 or 10x10) to hold the stats for each racer
        'if date, race#, track, and length are no longer the same as the last recorded, flush the array into the matrix

        If RaceDate = RacerCache(1, 1) And RaceNumber = RacerCache(1, 3) And RaceTrack = RacerCache(1, 4) And RaceLength = RacerCache(1, 5) Then
            'add racer to cache
            RacerCache(CacheInc, 1) = RaceDate
            RacerCache(CacheInc, 2) = RacerName
            RacerCache(CacheInc, 3) = RaceNumber
            RacerCache(CacheInc, 4) = RaceTrack
            RacerCache(CacheInc, 5) = RaceLength
            RacerCache(CacheInc, 6) = val1
            RacerCache(CacheInc, 7) = val2
            RacerCache(CacheInc, 8) = val3
            RacerCache(CacheInc, 9) = RaceStartBox
            RacerCache(CacheInc, 10) = RaceBreak
            RacerCache(CacheInc, 11) = RacerFinish
            RacerCache(CacheInc, 12) = RacerScore
            CacheInc = CacheInc + 1
        Else
            'if racer cache is empty THEN simply add racer to cache
            'else Flush racer cache to matrix, THEN add racer to cache
            If RacerCache(1, 1) Is Nothing Then
                RacerCache(CacheInc, 1) = RaceDate
                RacerCache(CacheInc, 2) = RacerName
                RacerCache(CacheInc, 3) = RaceNumber
                RacerCache(CacheInc, 4) = RaceTrack
                RacerCache(CacheInc, 5) = RaceLength
                RacerCache(CacheInc, 6) = val1
                RacerCache(CacheInc, 7) = val2
                RacerCache(CacheInc, 8) = val3
                RacerCache(CacheInc, 9) = RaceStartBox
                RacerCache(CacheInc, 10) = RaceBreak
                RacerCache(CacheInc, 11) = RacerFinish
                RacerCache(CacheInc, 12) = RacerScore
                CacheInc = CacheInc + 1
            Else
                If MatrixCalcMode = 0 Then
                    For av = 1 To 9
                        If RacerCache(av, 1) IsNot Nothing Then
                            Dim Racer1 As Integer
                            Dim Racer2 As Integer
                            Racer1 = RacerIDLookup(RacerCache(av, 2))
                            For au = av + 1 To 10
                                If RacerCache(au, 1) IsNot Nothing Then
                                    Racer2 = RacerIDLookup(RacerCache(au, 2))
                                    For aaq = 1 To 8
                                        If RacerMatrix(Racer1, Racer2, aaq, 1) = Nothing And HistoryInc = 0 Then
                                            HistoryInc = aaq
                                        End If
                                    Next aaq
                                    RacerMatrix(Racer1, Racer2, HistoryInc, 1) = CSng(RacerCache(av, 12) - RacerCache(au, 12))
                                    RacerMatrix(Racer2, Racer1, HistoryInc, 1) = CSng(RacerCache(au, 12) - RacerCache(av, 12))
                                    RacerMatrix(Racer1, Racer2, HistoryInc, 2) = CSng(RacerCache(av, 10) - RacerCache(au, 10))
                                    RacerMatrix(Racer2, Racer1, HistoryInc, 2) = CSng(RacerCache(au, 10) - RacerCache(av, 10))
                                    HistoryInc = 0
                                End If
                            Next au
                        End If
                        For aas = 1 To 12
                            RacerCache(av, aas) = Nothing
                        Next aas
                    Next av
                End If
                CacheInc = 1
                RacerCache(CacheInc, 1) = RaceDate
                RacerCache(CacheInc, 2) = RacerName
                RacerCache(CacheInc, 3) = RaceNumber
                RacerCache(CacheInc, 4) = RaceTrack
                RacerCache(CacheInc, 5) = RaceLength
                RacerCache(CacheInc, 6) = val1
                RacerCache(CacheInc, 7) = val2
                RacerCache(CacheInc, 8) = val3
                RacerCache(CacheInc, 9) = RaceStartBox
                RacerCache(CacheInc, 10) = RaceBreak
                RacerCache(CacheInc, 11) = RacerFinish
                RacerCache(CacheInc, 12) = RacerScore
                CacheInc = CacheInc + 1
            End If
        End If
    End Sub

    Private Function RacerIDLookup(ByVal lookupname As String)
        For at = 1 To RacerList.Count - 1
            If RacerList(at) = lookupname Then
                Return at
            End If
        Next at
    End Function

    Private Function RacerNameLookup(ByVal lookupid As Integer)
        Return RacerList(lookupid)
    End Function

    Private Sub ReadDataLines2(ByVal InputFile As String)
        If File.Exists(InputFile) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(InputFile)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & InputFile)
                Exit Sub
            End If
            Dim fs8 As New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr8 As New StreamReader(fs8)
            If _LastOffset8 < sr8.BaseStream.Length Then
                sr8.BaseStream.Seek(_LastOffset8, SeekOrigin.Begin)
                While sr8.Peek() <> -1
                    Dim read As String = sr8.ReadLine()
                    If read.Length <> 0 Then
                        PopulateRaceDataMatrix(read)
                    End If
                End While
                _LastOffset8 = sr8.BaseStream.Position
            End If
            sr8.Close()
            fs8.Close()
            _LastOffset8 = 0
            If MatrixCalcMode = 0 Then
                For av = 1 To 9
                    If RacerCache(av, 1) IsNot Nothing Then
                        Dim Racer1 As Integer
                        Dim Racer2 As Integer
                        Racer1 = RacerIDLookup(RacerCache(av, 2))
                        For au = av + 1 To 10
                            If RacerCache(au, 1) IsNot Nothing Then
                                Racer2 = RacerIDLookup(RacerCache(au, 2))
                                For aaq = 1 To 8
                                    If RacerMatrix(Racer1, Racer2, aaq, 1) = Nothing And HistoryInc = 0 Then
                                        HistoryInc = aaq
                                    End If
                                Next aaq
                                RacerMatrix(Racer1, Racer2, HistoryInc, 1) = CSng(RacerCache(av, 12) - RacerCache(au, 12))
                                RacerMatrix(Racer2, Racer1, HistoryInc, 1) = CSng(RacerCache(au, 12) - RacerCache(av, 12))
                                RacerMatrix(Racer1, Racer2, HistoryInc, 2) = CSng(RacerCache(av, 10) - RacerCache(au, 10))
                                RacerMatrix(Racer2, Racer1, HistoryInc, 2) = CSng(RacerCache(au, 10) - RacerCache(av, 10))
                                HistoryInc = 0
                            End If
                        Next au
                    End If
                    For aas = 1 To 12
                        RacerCache(av, aas) = Nothing 'purge cache
                    Next aas
                Next av
            End If
        End If
    End Sub

    Private Sub ConvertRacerBoxesToIDs()
        CollisionSubmittedNames = 0
        RecomputeIDs = False
        If DogBox1.Text IsNot "" Then
            If IsNumeric(DogBox1.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox1.Text))
                If aar IsNot Nothing Then
                    DogBox1.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox1.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox2.Text IsNot "" Then
            If IsNumeric(DogBox2.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox2.Text))
                If aar IsNot Nothing Then
                    DogBox2.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox2.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox3.Text IsNot "" Then
            If IsNumeric(DogBox3.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox3.Text))
                If aar IsNot Nothing Then
                    DogBox3.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox3.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox4.Text IsNot "" Then
            If IsNumeric(DogBox4.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox4.Text))
                If aar IsNot Nothing Then
                    DogBox4.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox4.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox5.Text IsNot "" Then
            If IsNumeric(DogBox5.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox5.Text))
                If aar IsNot Nothing Then
                    DogBox5.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox5.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox6.Text IsNot "" Then
            If IsNumeric(DogBox6.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox6.Text))
                If aar IsNot Nothing Then
                    DogBox6.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox6.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox7.Text IsNot "" Then
            If IsNumeric(DogBox7.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox7.Text))
                If aar IsNot Nothing Then
                    DogBox7.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox7.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
        If DogBox8.Text IsNot "" Then
            If IsNumeric(DogBox8.Text) = False Then
                Dim aar As String = ""
                aar = RacerIDLookup(PercentTwentyFilter(DogBox8.Text))
                If aar IsNot Nothing Then
                    DogBox8.Text = CStr(aar)
                    CollisionSubmittedNames = CollisionSubmittedNames + 1
                Else
                    NewEventMsg("No Racer found with name: " & DogBox8.Text & ".")
                    AbortCollision = 1
                    OmitNoDataRacer()
                End If
            Else
                CollisionSubmittedNames = CollisionSubmittedNames + 1
            End If
        End If
    End Sub

    Private Sub OmitNoDataRacer()
        If CollisionSubmittedNames = 0 Then
            DogBox1.Text = DogBox2.Text
            DogBox2.Text = DogBox3.Text
            DogBox3.Text = DogBox4.Text
            DogBox4.Text = DogBox5.Text
            DogBox5.Text = DogBox6.Text
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 1 Then
            DogBox2.Text = DogBox3.Text
            DogBox3.Text = DogBox4.Text
            DogBox4.Text = DogBox5.Text
            DogBox5.Text = DogBox6.Text
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 2 Then
            DogBox3.Text = DogBox4.Text
            DogBox4.Text = DogBox5.Text
            DogBox5.Text = DogBox6.Text
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 3 Then
            DogBox4.Text = DogBox5.Text
            DogBox5.Text = DogBox6.Text
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 4 Then
            DogBox5.Text = DogBox6.Text
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 5 Then
            DogBox6.Text = DogBox7.Text
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 6 Then
            DogBox7.Text = DogBox8.Text
            DogBox8.Text = ""
        End If
        If CollisionSubmittedNames = 7 Then
            DogBox8.Text = ""
        End If
        CollisionSubmittedNames = CollisionSubmittedNames + 1
        RecomputeIDs = True
    End Sub

    Private Function PercentTwentyFilter(ByVal inputtext As String)
        Dim returnstring As String = inputtext
        returnstring = returnstring.Replace(" ", "%20")
        Return returnstring
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AbortCollision = 0
        CreateCollision()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        'load a file in, read it in groups of eight, send eight racers to the collision handler and then save the output, repeat process
        OutputDataFormat = 1 ' always save in csv, do not save semantics in this way.
        BatchRacesProcessed = 0
        Dim BatchFileName As String = ""
        Dim RacersInCurrentGroup As Integer = 0
        OpenFileDialog1.ShowDialog()
        BatchFileName = OpenFileDialog1.FileName()
        NewEventMsg("Running batch collisions from file: " & BatchFileName)
        If File.Exists(BatchFileName) Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(BatchFileName)
            If fileinfo.Length < 1 Then
                NewEventMsg("ERROR: Supplied data file contains no data @ " & BatchFileName)
                Exit Sub
            End If
            Dim fs9 As New FileStream(BatchFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr9 As New StreamReader(fs9)
            If _LastOffset9 < sr9.BaseStream.Length Then
                sr9.BaseStream.Seek(_LastOffset9, SeekOrigin.Begin)
                While sr9.Peek() <> -1
                    Dim read As String = sr9.ReadLine()
                    If read.Length <> 0 Then
                        RacersInCurrentGroup = RacersInCurrentGroup + 1
                        AbortCollision = 0
                        If RacersInCurrentGroup = 1 Then
                            DogBox1.Text = read
                        End If
                        If RacersInCurrentGroup = 2 Then
                            DogBox2.Text = read
                        End If
                        If RacersInCurrentGroup = 3 Then
                            DogBox3.Text = read
                        End If
                        If RacersInCurrentGroup = 4 Then
                            DogBox4.Text = read
                        End If
                        If RacersInCurrentGroup = 5 Then
                            DogBox5.Text = read
                        End If
                        If RacersInCurrentGroup = 6 Then
                            DogBox6.Text = read
                        End If
                        If RacersInCurrentGroup = 7 Then
                            DogBox7.Text = read
                        End If
                        If RacersInCurrentGroup = 8 Then
                            DogBox8.Text = read
                            CollisionSubmittedNames = 8
                            CreateCollision()
                            If AbortCollision = 0 Then
                                Dim FileToCreate As String = ""
                                Dim currsecond As Integer = DateTime.Now.Second
                                Dim currmin As Integer = DateTime.Now.Minute
                                Dim currhour As Integer = DateTime.Now.Hour
                                Dim currday As Integer = DateTime.Now.Day
                                Dim currmon As Integer = DateTime.Now.Month
                                Dim curryr As Integer = DateTime.Now.Year
                                If My.Computer.FileSystem.DirectoryExists(ProgramDir & "\batch_races\") = False Then
                                    My.Computer.FileSystem.CreateDirectory(ProgramDir & "\batch_races\")
                                End If
                                FileToCreate = ProgramDir & "\batch_races\" & CStr(currday) & "-" & CStr(currmon) & "-" & CStr(curryr) & "\" & "Race" & CStr(BatchRacesProcessed) & "-" & CStr(currhour) & "-" & CStr(currmin) & "-" & CStr(currsecond) & ".csv"
                                If My.Computer.FileSystem.DirectoryExists(ProgramDir & "\batch_races\" & CStr(currday) & "-" & CStr(currmon) & "-" & CStr(curryr) & "\") = False Then
                                    My.Computer.FileSystem.CreateDirectory(ProgramDir & "\batch_races\" & CStr(currday) & "-" & CStr(currmon) & "-" & CStr(curryr) & "\")
                                End If
                                My.Computer.FileSystem.WriteAllText(FileToCreate, EventBox4.Text, True)
                            End If
                            If AbortCollision = 1 Then
                                Exit Sub
                            End If
                            ClearAllBoxes()
                            RacersInCurrentGroup = 0
                            BatchRacesProcessed = BatchRacesProcessed + 1
                        End If
                    End If
                End While
                _LastOffset9 = sr9.BaseStream.Position
            End If
            sr9.Close()
            fs9.Close()
            _LastOffset9 = 0
        End If
    End Sub

    Private Sub CreateCollision()
        ConvertRacerBoxesToIDs()
        While RecomputeIDs = True
            ConvertRacerBoxesToIDs() 'Prevents edge-cases
        End While
        If AbortCollision = 1 Then
            Return
        End If
        For ak = 1 To CollisionSubmittedNames - 1
            If ak = 1 Then
                CollisionX = CInt(DogBox1.Text)
            End If
            If ak = 2 Then
                CollisionX = CInt(DogBox2.Text)
            End If
            If ak = 3 Then
                CollisionX = CInt(DogBox3.Text)
            End If
            If ak = 4 Then
                CollisionX = CInt(DogBox4.Text)
            End If
            If ak = 5 Then
                CollisionX = CInt(DogBox5.Text)
            End If
            If ak = 6 Then
                CollisionX = CInt(DogBox6.Text)
            End If
            If ak = 7 Then
                CollisionX = CInt(DogBox7.Text)
            End If
            For aj = ak + 1 To CollisionSubmittedNames
                If aj = 2 Then
                    CollisionY = CInt(DogBox2.Text)
                End If
                If aj = 3 Then
                    CollisionY = CInt(DogBox3.Text)
                End If
                If aj = 4 Then
                    CollisionY = CInt(DogBox4.Text)
                End If
                If aj = 5 Then
                    CollisionY = CInt(DogBox5.Text)
                End If
                If aj = 6 Then
                    CollisionY = CInt(DogBox6.Text)
                End If
                If aj = 7 Then
                    CollisionY = CInt(DogBox7.Text)
                End If
                If aj = 8 Then
                    CollisionY = CInt(DogBox8.Text)
                End If
                For ap = 1 To 8 'first-gen search for immediately held data, as shown in file
                    Dim OutputValue As Single = 0
                    Dim OutputValue2 As Single = 0
                    OutputValue = RacerMatrix(CollisionX, CollisionY, ap, 1)
                    OutputValue2 = RacerMatrix(CollisionX, CollisionY, ap, 2)
                    If OutputValue = 0 Then
                        If ap = 1 Then
                            NewEventMsg2("No Data for racers: " & CStr(CollisionX) & " and " & CStr(CollisionY) & ".")
                        End If
                    Else
                        If OutputDataFormat = 0 Then
                            NewEventMsg2("Collision - Race: " & CStr(ap) & " for racers: " & CStr(CollisionX) & " and " & CStr(CollisionY) & " - Value: " & CStr(OutputValue) & "," & CStr(OutputValue2))
                        Else
                            NewEventMsg2(CStr(ap) & "," & CStr(CollisionX) & "," & CStr(CollisionY) & "," & CStr(OutputValue) & "," & CStr(OutputValue2))
                        End If
                    End If
                    UpdateBenchmark(1)
                Next ap
                Dim RollingAvg1 As Single = 0
                Dim RollingAvg1Break As Single = 0
                Dim Avg1contributions As Integer = 0
                Dim RollingAvg2 As Single = 0
                Dim RollingAvg2Break As Single = 0
                Dim Avg2contributions As Integer = 0
                Avg1contributions = 0
                Avg2contributions = 0
                RollingAvg1 = 0
                RollingAvg2 = 0
                RollingAvg1Break = 0
                RollingAvg2Break = 0
            Next aj
        Next ak
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim aar As String = ""
        aar = RacerIDLookup(TextBox3.Text)
        If aar IsNot Nothing Then
            NewEventMsg2("ID for racer " & TextBox3.Text & " Is:  " & CStr(aar))
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.Text = "CSV Format Data"
            OutputDataFormat = 1
        Else
            CheckBox1.Text = "Data w/ Semantics"
            OutputDataFormat = 0
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim aam As String = ""
        aam = RacerNameLookup(Val(TextBox3.Text))
        If aam = Nothing Then
            'skip
        Else
            NewEventMsg2("Racer ID " & TextBox3.Text & " belongs to: " & aam)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ClearAllBoxes()
    End Sub

    Private Sub ClearAllBoxes()
        DogBox1.Text = ""
        DogBox2.Text = ""
        DogBox3.Text = ""
        DogBox4.Text = ""
        DogBox5.Text = ""
        DogBox6.Text = ""
        DogBox7.Text = ""
        DogBox8.Text = ""
        EventBox.Text = ""
        EventBox2.Text = ""
        EventBox3.Text = ""
        EventBox4.Text = ""
    End Sub

    Private Sub CreateImageGraphFromMatrix()
        Dim bmp As New Bitmap(UniqueRacerNamesFound, UniqueRacerNamesFound)
        PictureBox1.Image = bmp
        Using g As Graphics = Graphics.FromImage(PictureBox1.Image)
            For ah = 0 To UniqueRacerNamesFound - 1
                For ai = 0 To UniqueRacerNamesFound - 1
                    If RacerMatrix(ah, ai, 1, 1) = Nothing Then
                        'skip pixel
                    Else
                        If RacerMatrix(ah, ai, 2, 1) = Nothing Then
                            bmp.SetPixel(ah, ai, Color.Red)
                        Else
                            If RacerMatrix(ah, ai, 3, 1) = Nothing Then
                                bmp.SetPixel(ah, ai, Color.Orange)
                            Else
                                If RacerMatrix(ah, ai, 4, 1) = Nothing Then
                                    bmp.SetPixel(ah, ai, Color.Yellow)
                                Else
                                    If RacerMatrix(ah, ai, 5, 1) = Nothing Then
                                        bmp.SetPixel(ah, ai, Color.Green)
                                    Else
                                        If RacerMatrix(ah, ai, 6, 1) = Nothing Then
                                            bmp.SetPixel(ah, ai, Color.Aqua)
                                        Else
                                            If RacerMatrix(ah, ai, 2, 1) = Nothing Then
                                                bmp.SetPixel(ah, ai, Color.Blue)
                                            Else
                                                If RacerMatrix(ah, ai, 8, 1) = Nothing Then
                                                    bmp.SetPixel(ah, ai, Color.Purple)
                                                Else
                                                    bmp.SetPixel(ah, ai, Color.HotPink)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ai
            Next ah
        End Using
        PictureBox1.Invalidate()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        CreateImageGraphFromMatrix()
        SaveFileDialog2.Title() = "Save Image As..."
        SaveFileDialog2.ShowDialog()
    End Sub

    Private Sub SaveFileDialog2_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog2.FileOk
        PictureBox1.Image.Save(SaveFileDialog2.FileName, System.Drawing.Imaging.ImageFormat.Png)
    End Sub

    '#################### NEW UI FUNCTIONS ####################

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button5.ForeColor = Color.Orange
        Button8.ForeColor = Color.Silver
        EventBox.Visible = True
        EventBox.Enabled = True
        EventBox2.Visible = False
        EventBox2.Enabled = False
        EventBox3.Visible = False
        EventBox3.Enabled = False
        EventBox4.Visible = False
        EventBox4.Enabled = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Button5.ForeColor = Color.Silver
        Button8.ForeColor = Color.Orange
        EventBox.Visible = False
        EventBox.Enabled = False
        EventBox2.Visible = True
        EventBox2.Enabled = True
        EventBox3.Visible = False
        EventBox3.Enabled = False
        EventBox4.Visible = False
        EventBox4.Enabled = False
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        Button5.ForeColor = Color.Silver
        Button8.ForeColor = Color.Silver
        EventBox.Visible = False
        EventBox.Enabled = False
        EventBox2.Visible = False
        EventBox2.Enabled = False
        EventBox3.Visible = True
        EventBox3.Enabled = True
        EventBox4.Visible = False
        EventBox4.Enabled = False
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)
        Button5.ForeColor = Color.Silver
        Button8.ForeColor = Color.Silver
        EventBox.Visible = False
        EventBox.Enabled = False
        EventBox2.Visible = False
        EventBox2.Enabled = False
        EventBox3.Visible = False
        EventBox3.Enabled = False
        EventBox4.Visible = True
        EventBox4.Enabled = True
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        LegacyFunctionsPanel.Enabled = True
        LegacyFunctionsPanel.Visible = True
        AnalysisNewFunctionsPanel.Enabled = False
        AnalysisNewFunctionsPanel.Visible = False
        ScraperPanel.Enabled = False
        ScraperPanel.Visible = False
        Button12.ForeColor = Color.Orange
        Button11.ForeColor = Color.Silver
        Button20.ForeColor = Color.Silver
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        LegacyFunctionsPanel.Enabled = False
        LegacyFunctionsPanel.Visible = False
        AnalysisNewFunctionsPanel.Enabled = True
        AnalysisNewFunctionsPanel.Visible = True
        ScraperPanel.Enabled = False
        ScraperPanel.Visible = False
        Button12.ForeColor = Color.Silver
        Button11.ForeColor = Color.Orange
        Button20.ForeColor = Color.Silver
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Close()
        Me.Dispose()
        Application.Exit()
    End Sub

    Private Sub TitleBar_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles AnalysisPanelTitleBar.MouseUp
        Go = False
        LeftSet = False
        TopSet = False
    End Sub

    Private Sub TitleBar_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles AnalysisPanelTitleBar.MouseDown
        Go = True
    End Sub

    Private Sub TitleBar_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles AnalysisPanelTitleBar.MouseMove
        If Go = True Then
            HoldLeft = (Control.MousePosition.X - Me.Left)
            HoldTop = (Control.MousePosition.Y - Me.Top)
            If TopSet = False Then
                OffTop = HoldTop - sender.Parent.Top
                TopSet = True
            End If
            If LeftSet = False Then
                OffLeft = HoldLeft - sender.Parent.Left
                LeftSet = True
            End If
            sender.Parent.Left = HoldLeft - OffLeft
            sender.Parent.Top = HoldTop - OffTop
            sender.Parent.Parent.Refresh()
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Me.TopMost = True Then
            Me.TopMost = False
        Else
            Me.TopMost = True
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim area As Rectangle
        Dim capture As Bitmap
        Dim graph As Graphics
        area = Me.Bounds
        capture = New Bitmap(area.Width, area.Height, Imaging.PixelFormat.Format32bppArgb)
        graph = Graphics.FromImage(capture)
        Me.Opacity = 0
        graph.CopyFromScreen(area.X, area.Y, 0, 0, area.Size, CopyPixelOperation.SourceCopy)
        Me.Opacity = OrigAlpha
        BackgroundRenderBoxMain.Image = capture
        graph.Dispose()
        ImageSelectionMode = 1
        Button15.ForeColor = Color.Orange
    End Sub

    Private Sub ImageSelect_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles BackgroundRenderBoxMain.MouseUp
        If ImageSelectionMode = 1 Then
            Dim area As Rectangle
            Dim capture As Bitmap
            Dim graph As Graphics
            ImageSelectX2 = MousePosition.X
            ImageSelectY2 = MousePosition.Y
            'capture image from area, save image to disk, then point PyTesseract to saved image.
            'Retrieve pyTesseracts result through console, and use it to populate the boxes if it fits the expected format
            NewEventMsg("Selected Image: " & CStr(ImageSelectX1) & ", " & CStr(ImageSelectY1) & " - " & CStr(ImageSelectX2) & ", " & CStr(ImageSelectY2))
            area = New Rectangle(x:=ImageSelectX1, y:=ImageSelectY1, width:=ImageSelectX2 - ImageSelectX1, height:=ImageSelectY2 - ImageSelectY1)
            capture = New Bitmap(ImageSelectX2 - ImageSelectX1, ImageSelectY2 - ImageSelectY1, Imaging.PixelFormat.Format32bppArgb)
            graph = Graphics.FromImage(capture)
            Me.Opacity = 0
            graph.CopyFromScreen(area.X, area.Y, 0, 0, area.Size, CopyPixelOperation.SourceCopy)
            If My.Computer.FileSystem.FileExists(ProgramDir & "readtext.png") = True Then
                My.Computer.FileSystem.DeleteFile(ProgramDir & "readtext.png")
            End If
            capture.Save(ProgramDir & "readtext.png", System.Drawing.Imaging.ImageFormat.Png)
            Me.Opacity = OrigAlpha
            graph.Dispose()
            BackgroundRenderBoxMain.Image = Nothing
            ReadImageTextWithTesseract(ProgramDir & "readtext.png")
            ImageSelectionMode = 0
            Button15.ForeColor = Color.Silver
        End If
    End Sub

    Private Sub ImageSelect_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles BackgroundRenderBoxMain.MouseDown
        If ImageSelectionMode = 1 Then
            ImageSelectX1 = MousePosition.X
            ImageSelectY1 = MousePosition.Y
        End If
    End Sub

    Private Sub ReadImageTextWithTesseract(SavedImagePath As String)
        If SavedImagePath IsNot "" Then
            Dim startinfo As ProcessStartInfo = New ProcessStartInfo()
            Dim proc As Process = New Process
            If My.Computer.FileSystem.DirectoryExists("E:\Programs\Anaconda\") = True Then
                startinfo.WorkingDirectory = My.Computer.FileSystem.CurrentDirectory + "\PyTesseract\"
                startinfo.FileName = "E:\Programs\Anaconda\python.exe"
                startinfo.Arguments = "TesseractReadFromImage.py " & ProgramDir & "readtext.png"
            Else
                If My.Computer.FileSystem.DirectoryExists("C:\Users\USER 2\") = True Then
                    startinfo.WorkingDirectory = Directory.GetCurrentDirectory + "\PyTesseract\"
                    startinfo.FileName = """C:\Users\USER 2\anaconda3\python.exe"""
                    startinfo.Arguments = "TesseractReadFromImage.py """ & ProgramDir & "readtext.png"""
                End If
            End If
            startinfo.UseShellExecute = False 'required for redirect.
            startinfo.WindowStyle = ProcessWindowStyle.Hidden 'dont show cmd window.
            startinfo.CreateNoWindow = True
            startinfo.RedirectStandardOutput = True 'capture output from cmd.
            proc.StartInfo = startinfo
            NewEventMsg(proc.StartInfo.WorkingDirectory)
            NewEventMsg(proc.StartInfo.FileName)
            NewEventMsg(proc.StartInfo.Arguments)
            Try
                proc.Start()
            Catch ex As Exception
                NewEventMsg(ex.Message)
            End Try
            Dim NumLines As Integer = 1
            Using procStreamReader As System.IO.StreamReader = proc.StandardOutput
                While procStreamReader.Peek() <> -1
                    Dim read As String = procStreamReader.ReadLine()
                    If read.Length <> 0 Then
                        If read.Length > 3 Then
                            If NumLines = 1 Then
                                DogBox1.Text = read
                            End If
                            If NumLines = 2 Then
                                DogBox2.Text = read
                            End If
                            If NumLines = 3 Then
                                DogBox3.Text = read
                            End If
                            If NumLines = 4 Then
                                DogBox4.Text = read
                            End If
                            If NumLines = 5 Then
                                DogBox5.Text = read
                            End If
                            If NumLines = 6 Then
                                DogBox6.Text = read
                            End If
                            If NumLines = 7 Then
                                DogBox7.Text = read
                            End If
                            If NumLines = 8 Then
                                DogBox8.Text = read
                            End If
                            NumLines = NumLines + 1
                        End If
                        If NumLines >= 9 Then
                            NewEventMsg(read)
                        End If
                        Application.DoEvents()
                    End If
                End While
            End Using
            proc.WaitForExit()
            proc.Dispose()
        Else
        End If
    End Sub

    '#################### IMPORTED SCRAPER FUNCTIONS ####################

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Button12.ForeColor = Color.Silver
        Button11.ForeColor = Color.Silver
        Button20.ForeColor = Color.Orange
        ScraperPanel.Enabled = True
        ScraperPanel.Visible = True
        LegacyFunctionsPanel.Enabled = False
        LegacyFunctionsPanel.Visible = False
        AnalysisNewFunctionsPanel.Enabled = False
        AnalysisNewFunctionsPanel.Visible = False
        'This is executed every time the scraper loads. We must check if our necessary file structure exists, and if not, create it.
        If Directory.Exists(ProgramLogDir) = False Then
            Directory.CreateDirectory(ProgramLogDir)
            Directory.CreateDirectory(ProgramCfgDir)
            Directory.CreateDirectory(ProgramLogDir)
            Directory.CreateDirectory(ProgramScrapeDir)
            File.Create(ProgramCfgDir & "Main.txt")
            NewEventMsg("Default Directory Structure was NOT found on the System @ " & ProgramDir)
            NewEventMsg("The default directory structure has been created. Populate CFG files first!")
        End If
        'If no errors occurred during startup, we should simply inform the user we are ready to process.
        NewEventMsg("Scraper filesystem exists. Ready!")
    End Sub

    Private Sub WriteEventLogMsg(ByRef LogMsg As String)
        'This makes outputting to the error log system easier in code.
        'Significantly more complex, because we want to record more data here than just program errors.
        'We will also record some statistics as well as stats from BenchMark here so we need to be able
        'to pass any string, and add a timestamp. It will be called by other subs to write data to the log file.
        Dim CurrLogFile As String = ProgramLogDir & DateAndTime.Now.Year & "\" & DateAndTime.Now.Month & "\" & DateAndTime.Now.Day & ".txt"
        If File.Exists(CurrLogFile) = False Then
            If Directory.Exists(ProgramLogDir & DateAndTime.Now.Year & "\" & DateAndTime.Now.Month & "\") = False Then
                Directory.CreateDirectory(ProgramLogDir & DateAndTime.Now.Year & "\" & DateAndTime.Now.Month & "\")
                NewEventMsg("Created Event Log Directory for current month.")
            End If
            NewEventMsg("Created Event Log File for current day.")
            'File.CreateText(CurrLogFile)
            Using sw As StreamWriter = File.AppendText(CurrLogFile)
                sw.WriteLine("[" & DateAndTime.Now.Hour & ":" & DateAndTime.Now.Minute & ":" & DateAndTime.Now.Second & "] -> " & LogMsg)
            End Using
            NewEventMsg(LogMsg)
        Else
            Using sw As StreamWriter = File.AppendText(CurrLogFile)
                sw.WriteLine("[" & DateAndTime.Now.Hour & ":" & DateAndTime.Now.Minute & ":" & DateAndTime.Now.Second & "] -> " & LogMsg)
            End Using
            NewEventMsg(LogMsg)
        End If
    End Sub

    Private Sub ReadCFGFile(ByVal CfgFile As String)
        If File.Exists(CfgFile) And EngineActive = 1 Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(CfgFile)
            If fileinfo.Length < 1 Then
                WriteEventLogMsg("ERROR: CFG File contains no data @ " & CfgFile)
                Exit Sub
            End If
            Dim CFGLineNum As Integer = 0
            Dim fs12 As New FileStream(CfgFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr12 As New StreamReader(fs12)
            If _ScrapeLastOffset < sr12.BaseStream.Length Then
                sr12.BaseStream.Seek(_ScrapeLastOffset, SeekOrigin.Begin)
                While sr12.Peek() <> -1
                    Dim read As String = sr12.ReadLine()
                    If read.Length <> 0 Then
                        ReadAppendFile(read, CFGLineNum)
                    End If
                    CFGLineNum = CFGLineNum + 1
                End While
                _ScrapeLastOffset = sr12.BaseStream.Position
            End If
            sr12.Close()
            fs12.Close()
            WriteEventLogMsg("FINISH: Scrape completed.")
            WriteEventLogMsg("Iterated over " & CFGLineNum & " addresses, Scanned " & PagesScanned & " Web Pages, and Archived " & PagesArchived & " of them.")
            CFGLineNum = 0
            PagesScanned = 0
            PagesArchived = 0
            _ScrapeLastOffset = 0
        End If
    End Sub

    Private Sub ReadAppendFile(ByVal CfgLineRead As String, ByVal CfgLineNum As Integer)
        Dim AppendDir As String = ProgramCfgDir & CfgLineNum & "\"
        Dim AppendFile As String = "append.txt"
        Dim AppendLine As Integer = 0
        If File.Exists(AppendDir + AppendFile) = True Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(AppendDir + AppendFile)
            If fileinfo.Length < 1 Then
                WriteEventLogMsg("ERROR: CFG-Appenditure File contains no data @ " & AppendDir + AppendFile)
                Exit Sub
            End If
            Dim fs13 As New FileStream(AppendDir + AppendFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr13 As New StreamReader(fs13)
            If EngineStopped = 1 Then
                WriteEventLogMsg("STOP: Scraping operation was stopped.")
                Exit Sub
            End If
            If EngineActive = 1 Then
                If _ScrapeLastOffset2 < sr13.BaseStream.Length Then
                    sr13.BaseStream.Seek(_ScrapeLastOffset2, SeekOrigin.Begin)
                    While sr13.Peek() <> -1
                        Dim read As String = sr13.ReadLine()
                        If read.Length <> 0 Then
                            AssembleURL(CfgLineRead, CfgLineNum, read, AppendLine)
                            Application.DoEvents()
                        End If
                        AppendLine = AppendLine + 1
                        If EngineStopped = 1 Then
                            WriteEventLogMsg("STOP: Scraping operation was stopped.")
                            Exit While
                        End If
                    End While
                    _ScrapeLastOffset2 = sr13.BaseStream.Position
                End If
            End If
            sr13.Close()
            fs13.Close()
            AppendLine = 0
            _ScrapeLastOffset2 = 0
        End If
        If File.Exists(AppendDir + AppendFile) = False Then
            If Directory.Exists(AppendDir) = False Then
                Directory.CreateDirectory(AppendDir)
                WriteEventLogMsg("Created Appenditure Directory for current CFG line# " & CStr(CfgLineNum))
            End If
            File.Create(AppendDir + AppendFile)
            WriteEventLogMsg("Created Appenditure File for CFG-file line# " & CStr(CfgLineNum))
        End If
    End Sub

    Private Sub AssembleURL(ByVal CfgLineRead As String, ByVal CfgLineNum As Integer, ByVal AppendLineRead As String, ByVal AppendLineNum As Integer)
        PagesScanned = PagesScanned + 1
        Dim ArchiveFileName As String = ProgramArchiveDir & AppendLineRead & ".txt"
        If ErrorCount > ErrorCountMax Then
            WriteEventLogMsg("WARNING: Recieved an error code from requested address " & ErrorCount & " Times, Scrape operation cancelled.")
            EngineActive = 0
            EngineStopped = 1
            Exit Sub
        End If
        If Directory.Exists(ProgramArchiveDir) = True Then
            If File.Exists(ArchiveFileName) = True Then
                Dim HoldString As String = File.ReadAllText(ArchiveFileName)
                Dim MostRecentDate As String = HoldString.Remove(HoldString.IndexOf(","), HoldString.Length - HoldString.IndexOf(","))
                Try
                    Dim sourceString As String = New System.Net.WebClient().DownloadString(CfgLineRead & AppendLineRead)
                    sourceString = ParseFileData(sourceString, AppendLineRead)
                    sourceString = sourceString.Remove(sourceString.IndexOf(MostRecentDate), sourceString.Length() - sourceString.IndexOf(MostRecentDate))
                    File.WriteAllText(ArchiveFileName, sourceString & HoldString)
                    WriteEventLogMsg("Address: " & CfgLineRead & AppendLineRead & " --Archived to: " & ArchiveFileName)
                    PagesArchived = PagesArchived + 1
                Catch ex As Exception
                    WriteEventLogMsg("Recieved Error Message: " & ex.Message & " -- Retrying...")
                    Try
                        Dim sourceString As String = New System.Net.WebClient().DownloadString(CfgLineRead & AppendLineRead)
                        sourceString = ParseFileData(sourceString, AppendLineRead)
                        sourceString = sourceString.Remove(sourceString.IndexOf(MostRecentDate), sourceString.Length() - sourceString.IndexOf(MostRecentDate))
                        File.WriteAllText(ArchiveFileName, sourceString & HoldString)
                        WriteEventLogMsg("Address: " & CfgLineRead & AppendLineRead & " --Archived to: " & ArchiveFileName)
                        PagesArchived = PagesArchived + 1
                    Catch ex2 As Exception
                        WriteEventLogMsg("Recieved Error Message: " & ex.Message)
                        ErrorCount = ErrorCount + 1
                    End Try
                End Try
            End If
            If File.Exists(ArchiveFileName) = False Then
                Try
                    Dim sourceString As String = New System.Net.WebClient().DownloadString(CfgLineRead & AppendLineRead)
                    sourceString = ParseFileData(sourceString, AppendLineRead)
                    File.WriteAllText(ArchiveFileName, sourceString)
                    WriteEventLogMsg("Address: " & CfgLineRead & AppendLineRead & " --Archived to: " & ArchiveFileName)
                    PagesArchived = PagesArchived + 1
                Catch ex As Exception
                    WriteEventLogMsg("Recieved Error Message: " & ex.Message & " -- Retrying...")
                    Try
                        Dim sourceString As String = New System.Net.WebClient().DownloadString(CfgLineRead & AppendLineRead)
                        sourceString = ParseFileData(sourceString, AppendLineRead)
                        File.WriteAllText(ArchiveFileName, sourceString)
                        WriteEventLogMsg("Address: " & CfgLineRead & AppendLineRead & " --Archived to: " & ArchiveFileName)
                        PagesArchived = PagesArchived + 1
                    Catch ex2 As Exception
                        WriteEventLogMsg("Recieved Error Message: " & ex.Message)
                        ErrorCount = ErrorCount + 1
                    End Try
                End Try
            End If
        End If
    End Sub

    'DB Migration Functions
    Private Function ParseFileData(ByVal inputstring As String, ByVal appendread As String)
        If inputstring.Length > 0 Then
            Dim ParseData As String = inputstring
            Dim matchpattern0 As String = "<.*""raceline"" nowrap>[^<]*</td>"
            Dim matchpattern1 As String = "td\.raceline.?\{.*\}"
            Dim matchpattern2 As String = "<(?:[^>=]|='[^']*'|=""[^""]*""|=[^'""][^\s>]*)*>"
            Dim matchpattern3 As String = "\s+"
            Dim matchpattern4 As String = ",\s*(?=\d{4}-\d{2}-\d{2})"
            Dim matchpattern5 As String = "(?<date>\d{4}-\d{2}-\d{2}),"
            ParseData = Regex.Replace(ParseData, matchpattern0, "")
            ParseData = Regex.Replace(ParseData, matchpattern1, "")
            ParseData = Regex.Replace(ParseData, matchpattern2, "")
            ParseData = WebUtility.HtmlDecode(ParseData).Trim()
            ParseData = Regex.Replace(ParseData, matchpattern3, ",")
            ParseData = Regex.Replace(ParseData, matchpattern4, Environment.NewLine)
            ParseData = Regex.Replace(ParseData, matchpattern5, "${date}," & appendread & ",")
            Return ParseData
        End If
    End Function

    Private Sub ConcatFiles()
        Dim ConCatFile As String = ProgramCfgDir & "Concat.txt"
        Dim ConCatOutput As String = ProgramScrapeDir & "ConCat-" & DateAndTime.Now.Year & "-" & DateAndTime.Now.Month & "-" & DateAndTime.Now.Day & ".txt"
        Dim ConCatString As String = ""
        Dim Iterations As Integer = 0
        WriteEventLogMsg("Concatenating archived information to single-file...")
        If File.Exists(ConCatFile) = True Then
            Dim fileinfo As FileInfo = My.Computer.FileSystem.GetFileInfo(ConCatFile)
            If fileinfo.Length < 1 Then
                WriteEventLogMsg("ERROR: Concatenation Info File contains no data @ " & ConCatFile)
                Exit Sub
            End If
            Dim fs14 As New FileStream(ConCatFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Dim sr14 As New StreamReader(fs14)
            If _ScrapeLastOffset3 < sr14.BaseStream.Length Then
                sr14.BaseStream.Seek(_ScrapeLastOffset3, SeekOrigin.Begin)
                While sr14.Peek() <> -1
                    Dim read As String = sr14.ReadLine()
                    If read.Length <> 0 Then
                        Dim FileToConcat As String = ProgramArchiveDir & read & ".txt"
                        If FileToConcat.Contains(vbTab) = True Then
                            FileToConcat = FileToConcat.Remove(FileToConcat.IndexOf(vbTab), FileToConcat.Length - FileToConcat.IndexOf(vbTab)) & ".txt"
                        End If
                        Dim FileLines As String() = File.ReadAllLines(FileToConcat)
                        Dim ConcatSR As StreamReader = File.OpenText(FileToConcat)
                        If FileLines.Length > ConcatMaxLines Then
                            For a = 1 To ConcatMaxLines
                                ConCatString = ConCatString & ConcatSR.ReadLine() & Environment.NewLine
                            Next a
                        End If
                        If FileLines.Length <= ConcatMaxLines Then
                            While ConcatSR.Peek() <> -1
                                ConCatString = ConCatString & ConcatSR.ReadLine() & Environment.NewLine
                            End While
                        End If
                        Iterations = Iterations + 1
                        Application.DoEvents()
                    End If
                End While
                _ScrapeLastOffset3 = sr14.BaseStream.Position
            End If
            sr14.Close()
            fs14.Close()
            _ScrapeLastOffset3 = 0
            Dim matchpattern As String = Environment.NewLine & Environment.NewLine
            ConCatString = Regex.Replace(ConCatString, matchpattern, Environment.NewLine)
            File.WriteAllText(ConCatOutput, ConCatString)
            WriteEventLogMsg("Concatenated up to " & ConcatMaxLines & " lines of data from " & Iterations & " selected sources.")
            Iterations = 0
        End If
    End Sub

    'UI Elements
    Private Sub ScapeStartButton_Click(sender As Object, e As EventArgs) Handles ScapeStartButton.Click
        EngineActive = 1
        EngineStopped = 0
        EnginePaused = 0
        WriteEventLogMsg("START: Scrape has started...")
        ReadCFGFile(ProgramCfgDir & "Main.txt")
    End Sub

    Private Sub ScrapePauseButton_Click(sender As Object, e As EventArgs) Handles ScrapePauseButton.Click
        If EnginePaused = 1 Then
            EnginePaused = 0
        Else
            EnginePaused = 1
        End If
    End Sub

    Private Sub ScrapeStopButton_Click(sender As Object, e As EventArgs) Handles ScrapeStopButton.Click
        EngineActive = 0
        EngineStopped = 1
    End Sub

    Private Sub ScrapeConcatButton_Click(sender As Object, e As EventArgs) Handles ScrapeConcatButton.Click
        ConcatFiles()
    End Sub

    Private Sub ScrapeErrorSelect_ValueChanged(sender As Object, e As EventArgs) Handles ScrapeErrorSelect.ValueChanged
        ErrorCountMax = ScrapeErrorSelect.Value - 1
    End Sub

    '#################### BENCHMARKING ####################

    Private Sub BenchmarkUtilityInitialize()

    End Sub

    Private Sub UpdateBenchmark(stage As Integer)
        If BenchmarkingMode = 1 Then
            BenchmarkCurrTime = DateTime.Now.Millisecond
            BenchmarkLastIterationTime = BenchmarkCurrTime - BenchmarkPrevTime
            If BenchmarkCurrentIteration < 10 Then
                BenchmarkIterationData(BenchmarkCurrentIteration) = BenchmarkLastIterationTime
                BenchmarkCurrentIteration = BenchmarkCurrentIteration + 1
                BenchmarkIterationsThisSecond = BenchmarkIterationsThisSecond + 1
            ElseIf BenchmarkCurrentIteration = 10 Then
                BenchmarkIterationData(BenchmarkCurrentIteration) = BenchmarkLastIterationTime
                If BenchmarkCurrentAvgIteration < 10 Then
                    BenchmarkIterationAverages(BenchmarkCurrentAvgIteration) = (BenchmarkIterationData(1) + BenchmarkIterationData(2) + BenchmarkIterationData(3) + BenchmarkIterationData(4) + BenchmarkIterationData(5) + BenchmarkIterationData(6) + BenchmarkIterationData(7) + BenchmarkIterationData(8) + BenchmarkIterationData(9) + BenchmarkIterationData(10)) / 10
                    BenchmarkCurrentAvgIteration = BenchmarkCurrentAvgIteration + 1
                ElseIf BenchmarkCurrentAvgIteration = 10 Then
                    BenchmarkIterationAverages(BenchmarkCurrentAvgIteration) = (BenchmarkIterationData(1) + BenchmarkIterationData(2) + BenchmarkIterationData(3) + BenchmarkIterationData(4) + BenchmarkIterationData(5) + BenchmarkIterationData(6) + BenchmarkIterationData(7) + BenchmarkIterationData(8) + BenchmarkIterationData(9) + BenchmarkIterationData(10)) / 10
                    BenchmarkCurrentAvgIteration = 1
                End If
                BenchmarkCurrentIteration = 1
                BenchmarkIterationsThisSecond = BenchmarkIterationsThisSecond + 1
            End If
            If BenchmarkCurrTime > BenchmarkNextUpdateTime Then
                BenchmarkStageLabel.Text = "Stage: " & CStr(stage)
                BenchmarkItsPerSecondLabel.Text = CStr(BenchmarkIterationsThisSecond) & " it/s"
                BenchmarkIterationTimeLabel.Text = CStr(BenchmarkIterationAverages(BenchmarkCurrentAvgIteration)) & " ms"
                BenchmarkNextUpdateTime = BenchmarkCurrTime + 1000
                BenchmarkIterationsThisSecond = 0
            End If
            If BenchmarkDoEvents = True Then
                Application.DoEvents()
            End If
        End If
    End Sub

    Private Sub BenchmarkStateCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles BenchmarkStateCheckbox.CheckedChanged
        If BenchmarkStateCheckbox.Checked = True Then
            BenchmarkStateCheckbox.ForeColor = Color.Green
            BenchmarkStateCheckbox.Text = "Enabled"
            BenchmarkingMode = 1
        Else
            BenchmarkStateCheckbox.ForeColor = Color.Red
            BenchmarkStateCheckbox.Text = "Disabled"
            BenchmarkingMode = 0
        End If
    End Sub

    Private Sub BenchmarkDoeventsCheckbox_CheckedChanged(sender As Object, e As EventArgs) Handles BenchmarkDoeventsCheckbox.CheckedChanged
        If BenchmarkDoeventsCheckbox.Checked = True Then
            BenchmarkDoeventsCheckbox.ForeColor = Color.Green
            BenchmarkDoeventsCheckbox.Text = "DoEvents ON"
            BenchmarkDoEvents = True
        Else
            BenchmarkDoeventsCheckbox.ForeColor = Color.Red
            BenchmarkDoeventsCheckbox.Text = "DoEvents OFF"
            BenchmarkDoEvents = False
        End If
    End Sub


    '#################### NEW STUFF ####################

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        For dg = 1 To RaceDataArraySize
            If RacerList(dg) IsNot Nothing Then
                NewEventMsg(RacerList(dg) & "," & CStr(dg))
            End If
        Next dg
    End Sub
End Class
