Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions

Public Module mdlSWMM
    'Public Const appPath As String = "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll"
    ' SWMM5_IFACE.BAS
    ' Example code for interfacing SWMM 5
    ' with Visual Basic Applications
    ' Remember to add swmm5.dll to the application start up path.
    ' Declarations of imported procedures from the EPASWMM DLL engine (SWMM5.DLL)
    ''Declare Function swmm_run Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Integer
    ''Declare Function swmm_open Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Integer
    ''Declare Function swmm_start Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByVal saveFlag As Integer) As Integer
    Declare Function swmm_run Lib "swmm5.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Integer
    Declare Function swmm_open Lib "swmm5.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Integer
    Declare Function swmm_start Lib "swmm5.dll" (ByVal saveFlag As Integer) As Integer
    '<DllImport("C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll", CallingConvention:=CallingConvention.Cdecl, CharSet:=CharSet.Ansi)>
    'Public Function swmm_step(elapsedTime As Single) As Integer
    'End Function
    ''Declare Function swmm_step Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByRef elapsedTime As Double) As Integer
    ''Declare Function swmm_end Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" () As Integer
    ''Declare Function swmm_report Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" () As Integer
    ''Declare Function swmm_getMassBalErr Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByRef runoffErr As Single, ByRef flowErr As Single, ByRef qualErr As Single) As Integer
    ''Declare Function swmm_close Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" () As Integer
    ''Declare Function swmm_getVersion Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" () As Integer
    ''Declare Function swmm_getWarnings Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" () As Integer
    ''Declare Function swmm_getError Lib "C:\Program Files (x86)\EPA SWMM 5.1.013\swmm5.dll" (ByVal errMsg As StringBuilder, ByVal msgLen As Integer) As Integer
    Declare Function swmm_step Lib "swmm5.dll" (ByRef elapsedTime As Double) As Integer
    Declare Function swmm_end Lib "swmm5.dll" () As Integer
    Declare Function swmm_report Lib "swmm5.dll" () As Integer
    Declare Function swmm_getMassBalErr Lib "swmm5.dll" (ByRef runoffErr As Single, ByRef flowErr As Single, ByRef qualErr As Single) As Integer
    Declare Function swmm_close Lib "swmm5.dll" () As Integer
    Declare Function swmm_getVersion Lib "swmm5.dll" () As Integer
    Declare Function swmm_getWarnings Lib "swmm5.dll" () As Integer
    Declare Function swmm_getError Lib "swmm5.dll" (ByVal errMsg As StringBuilder, ByVal msgLen As Integer) As Integer
    'Holding the binary file instance
    Public fsAs As FileStream
    Public fIn As BinaryReader

    Private Structure STARTUPINFO
        Public cb As Integer
        Public lpReserved As String
        Public lpDesktop As String
        Public lpTitle As String
        Public dwX As Integer
        Public dwY As Integer
        Public dwXSize As Integer
        Public dwYSize As Integer
        Public dwXCountChars As Integer
        Public dwYCountChars As Integer
        Public dwFillAttribute As Integer
        Public dwFlags As Integer
        Public wShowWindow As Integer
        Public cbReserved2 As Integer
        Public lpReserved2 As Integer
        Public hStdInput As Integer
        Public hStdOutput As Integer
        Public hStdError As Integer
    End Structure

    Private Structure PROCESS_INFORMATION
        Public hProcess As Integer
        Public hThread As Integer
        Public dwProcessID As Integer
        Public dwThreadID As Integer
    End Structure

    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
         hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer

    Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer,
      ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer,
      ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As String,
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Integer

    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer

    Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Integer, lpExitCode As Integer) As Integer

    Private Const SUBCATCH = 0
    Private Const NODE = 1
    Private Const LINK = 2
    Private Const SYS = 3
    Private Const INFINITE = -1&
    Private Const SW_SHOWNORMAL = 1&
    Private Const RECORDSIZE = 4           ' number of bytes per file record

    Private SubcatchVars As Integer           ' number of subcatch reporting variable
    Private NodeVars As Integer               ' number of node reporting variables
    Private LinkVars As Integer               ' number of link reporting variables
    Private SysVars As Integer                ' number of system reporting variables
    Private StartPos As Integer               ' file position where results start
    Private BytesPerPeriod As Integer         ' number of bytes used for storing
    ' results in file each reporting period
    Public SWMM_Nperiods As Integer           ' number of reporting periods
    Public SWMM_FlowUnits As Integer          ' flow units code
    Public SWMM_Nsubcatch As Integer          ' number of subcatchments
    Public SWMM_Nnodes As Integer             ' number of drainage system nodes
    Public SWMM_Nlinks As Integer             ' number of drainage system links
    Public SWMM_Npolluts As Integer           ' number of pollutants tracked
    Public SWMM_StartDate As Double        ' start date of simulation
    Public SWMM_ReportStep As Integer         ' reporting time step (seconds)

    Public Function RunSwmmExe(cmdLine As String) As Integer
        '------------------------------------------------------------------------------
        '  Input:   cmdLine = command line for running the console version of SWMM 5
        '  Output:  returns the exit code generated by running SWMM5.EXE
        '  Purpose: runs the command line version of SWMM 5.
        '------------------------------------------------------------------------------
        Dim pi As PROCESS_INFORMATION
        Dim si As STARTUPINFO
        Dim exitCode As Integer
        ' --- Initialize data structures
        si.cb = Len(si)
        si.wShowWindow = SW_SHOWNORMAL
        ' --- launch swmm5.exe
        exitCode = CreateProcessA(vbNullString, cmdLine, 0&, 0&, 0&,
                        0&, 0&, vbNullString, si, pi)
        ' --- wait for program to end
        exitCode = WaitForSingleObject(pi.hProcess, INFINITE)
        ' --- retrieve the error code produced by the program
        Call GetExitCodeProcess(pi.hProcess, exitCode)
        ' --- release handles
        Call CloseHandle(pi.hThread)
        Call CloseHandle(pi.hProcess)
        RunSwmmExe = exitCode
    End Function


    Public Function RunSwmmDll(inpFile As String, rptFile As String,
           OutFile As String) As Integer
        '------------------------------------------------------------------------------
        '  Input:   inpFile = name of SWMM 5 input file
        '           rptFile = name of status report file
        '           outFile = name of binary output file
        '  Output:  returns a SWMM 5 error code or 0 if there are no errors
        '  Purpose: runs the dynamic link library version of SWMM 5.
        '------------------------------------------------------------------------------
        Dim err As Integer, elapsedTime As Double
        ' --- open a SWMM project
        err = swmm_open(inpFile, rptFile, OutFile)
        If err = 0 Then
            ' --- initialize all processing systems
            err = swmm_start(1)
            If err = 0 Then
                ' --- step through the simulation
                Do
                    ' --- allow Windows to process any pending events
                    Application.DoEvents()
                    ' --- extend the simulation by one routing time step
                    err = swmm_step(elapsedTime)
                    '//////////////////////////////////////////
                    ' call a progress reporting function here,
                    ' using elapsedTime as an argument
                    '//////////////////////////////////////////
                Loop While elapsedTime > 0# And err = 0
            End If
            ' --- close all processing systems
            swmm_end
        End If
        ' --- close the project
        swmm_close
        ' --- return the error code
        RunSwmmDll = err
    End Function

    Function OpenSwmmOutFile(OutFile As String) As Integer
        '------------------------------------------------------------------------------
        '  Input:   outFile = name of binary output file
        '  Output:  returns 0 if successful, 1 if binary file invalid because
        '           SWMM 5 ran with errors, or 2 if the file cannot be opened
        '  Purpose: opens the binary output file created by a SWMM 5 run and
        '           retrieves the following simulation data that can be
        '           accessed by the application:
        '           SWMM_Nperiods = number of reporting periods
        '           SWMM_FlowUnits = flow units code
        '           SWMM_Nsubcatch = number of subcatchments
        '           SWMM_Nnodes = number of drainage system nodes
        '           SWMM_Nlinks = number of drainage system links
        '           SWMM_Npolluts = number of pollutants tracked
        '           SWMM_StartDate = start date of simulation
        '           SWMM_ReportStep = reporting time step (seconds)
        '------------------------------------------------------------------------------
        Dim magic1 As Integer, magic2 As Integer
        Dim errCode As Integer, version As Integer, offset As Integer
        Dim offset0 As Integer, err As Integer

        fsAs = New FileStream(OutFile, FileMode.Open, FileAccess.Read)
        fIn = New BinaryReader(fsAs)
        ' --- open the output file
        On Error GoTo FINISH
        err = 2
        ' --- check that file contains at least 14 records
        If fsAs.Length < 14 * RECORDSIZE Then Exit Function
        'Set the cursor position from end of the file. vb.net is 0 index based where as vb6 1 index based.
        fsAs.Position = fsAs.Length - 5 * RECORDSIZE 'vb6 add 1 more to the length
        'Debug.WriteLine(fsAs.Position & "-" & Len(offset0))
        ' --- read parameters from end of file
        offset0 = fIn.ReadInt32()
        StartPos = fIn.ReadInt32()
        SWMM_Nperiods = fIn.ReadInt32()
        errCode = fIn.ReadInt32()
        magic2 = fIn.ReadInt32()
        ' --- read magic number from beginning of file
        fsAs.Position = 0 'vb6- base index=1
        magic1 = fIn.ReadInt32()
        ' --- perform error checks
        If magic1 <> magic2 Then
            err = 1
        ElseIf errCode <> 0 Then
            err = 1
        ElseIf SWMM_Nperiods = 0 Then
            err = 1
        Else
            err = 0
        End If
        ' --- quit if errors found
        If err > 0 Then Exit Function
        ' --- otherwise read additional parameters from start of file
        version = fIn.ReadInt32()
        SWMM_FlowUnits = fIn.ReadInt32()
        SWMM_Nsubcatch = fIn.ReadInt32()
        SWMM_Nnodes = fIn.ReadInt32()
        SWMM_Nlinks = fIn.ReadInt32()
        SWMM_Npolluts = fIn.ReadInt32()
        ' --- skip over saved subcatch/node/link input values
        offset = (SWMM_Nsubcatch + 2) * RECORDSIZE
        offset = offset + (3 * SWMM_Nnodes + 4) * RECORDSIZE
        offset = offset + (5 * SWMM_Nlinks + 6) * RECORDSIZE
        fsAs.Position = offset0 + offset
        ' --- read number & codes of computed variables
        SubcatchVars = fIn.ReadInt32()
        fsAs.Position += (SubcatchVars * RECORDSIZE)
        NodeVars = fIn.ReadInt32()
        fsAs.Position += (NodeVars * RECORDSIZE)
        LinkVars = fIn.ReadInt32()
        fsAs.Position += (LinkVars * RECORDSIZE)
        SysVars = fIn.ReadInt32()
        ' --- read data just before start of output results
        fsAs.Position = StartPos - 3 * RECORDSIZE
        SWMM_StartDate = fIn.ReadDouble()
        SWMM_ReportStep = fIn.ReadInt32()
        ' --- compute number of bytes stored per reporting period
        BytesPerPeriod = RECORDSIZE * 2
        BytesPerPeriod = BytesPerPeriod + RECORDSIZE * SWMM_Nsubcatch * SubcatchVars
        BytesPerPeriod = BytesPerPeriod + RECORDSIZE * SWMM_Nnodes * NodeVars
        BytesPerPeriod = BytesPerPeriod + RECORDSIZE * SWMM_Nlinks * LinkVars
        BytesPerPeriod = BytesPerPeriod + RECORDSIZE * SysVars
        ' --- return with file left open
        OpenSwmmOutFile = err
        Exit Function
FINISH:
        OpenSwmmOutFile = err
        fIn.Close()
    End Function

    Function GetSwmmResult(ByVal iType As Integer, ByVal iIndex As Integer,
             ByVal vIndex As Integer, ByVal period As Integer, ByRef Value As Single) As Integer
        '------------------------------------------------------------------------------
        '  Input:   iType = type of object whose value is being sought
        '                   (0 = subcatchment, 1 = node, 2 = link, 3 = system
        '           iIndex = index of item being sought (starting from 0)
        '           vIndex = index of variable being sought (see Interfacing Guide)
        '           period = reporting period index (starting from 1)
        '  Output:  value = value of variable being sought;
        '           function returns 1 if successful, 0 if not
        '  Purpose: finds the result of a specific variable for a given object
        '           at a specified time period.
        '------------------------------------------------------------------------------
        Dim offset1 As Integer, offset2 As Integer
        '// --- compute offset into output file
        Value = 0#
        GetSwmmResult = 0
        offset1 = StartPos + (period - 1) * BytesPerPeriod + 2 * RECORDSIZE + 1
        offset2 = 0
        If iType = SUBCATCH Then
            offset2 = iIndex * SubcatchVars + vIndex
        ElseIf iType = NODE Then
            offset2 = SWMM_Nsubcatch * SubcatchVars + iIndex * NodeVars + vIndex
        ElseIf iType = LINK Then
            offset2 = SWMM_Nsubcatch * SubcatchVars + SWMM_Nnodes * NodeVars + iIndex * LinkVars + vIndex
        ElseIf iType = SYS Then
            offset2 = SWMM_Nsubcatch * SubcatchVars + SWMM_Nnodes * NodeVars + SWMM_Nlinks * LinkVars + vIndex
        Else : Exit Function
        End If
        '// --- re-position the file and read result
        fsAs.Position = offset1 + RECORDSIZE * offset2 - 1
        Value = fIn.ReadSingle()
        GetSwmmResult = 1
    End Function

    Public Sub CloseSwmmOutFile()
        '------------------------------------------------------------------------------
        '  Input:   none
        '  Output:  none
        '  Purpose: closes the binary output file.
        '------------------------------------------------------------------------------
        Try
            If (fIn Is Nothing) Then GoTo gt
            fIn.Close()
        Catch 'ex As Exception
            'do nothing
        End Try
gt:
        Try
            If (fsAs Is Nothing) Then Exit Sub
            fsAs.Close()
        Catch 'ex As Exception
            'do nothing
        End Try
    End Sub

    Public Function read_TimeSeriesQ(ByVal outFile As String,
                        conIndex As Integer, ByVal startDate As String) As List(Of String)
        Dim QtimeSeries As New List(Of String)
        Dim lngDate As Single
        If (outFile = String.Empty) Then Exit Function
        On Error GoTo ErrH
        Dim r As Long, i As Long, flowrate As Single
        r = OpenSwmmOutFile(outFile)

        lngDate = Convert.ToInt32(Convert.ToDateTime(startDate).ToOADate)
        Dim startPeriod As Long, endPeriod As Long, swmm_hrPeriod As Double
        swmm_hrPeriod = 60 / (SWMM_ReportStep / 60)  'Converting the step of seconds to numerical.
        startPeriod = (lngDate - SWMM_StartDate) * 24 * swmm_hrPeriod
        endPeriod = startPeriod + 24 * swmm_hrPeriod - 1
        For i = startPeriod To endPeriod
            GetSwmmResult(2, conIndex, 0, i, flowrate) 'Index of conduit element has to identified manually.
            Dim lngTimeStamp As Double = lngDate + ((i - startPeriod) / swmm_hrPeriod) / 24
            Dim txtTimeStamp As String = DateTime.FromOADate(lngTimeStamp).ToString()
            txtTimeStamp += "," + flowrate.ToString()
            QtimeSeries.Add(txtTimeStamp)
        Next
ErrH:
        On Error GoTo 0
        CloseSwmmOutFile()
        Return QtimeSeries
    End Function

    '  ********************
    'Link Flow Summary
    '********************

    '-----------------------------------------------------------------------------
    '                               Maximum  Time Of Max   Maximum    Max/    Max/
    '                                |Flow|   Occurrence   |Veloc|    Full    Full
    'Link                 Type          MGD  days hr:min    ft/sec    Flow   Depth
    '-----------------------------------------------------------------------------
    Public Function read_LinkFlowSummary(rptFile As String) As Dictionary(Of String, String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New Dictionary(Of String, String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Link Flow Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 7
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                i = 5
                Dim uID As String = cOut(0).Trim()
                Dim rslt As String = uID
                If (cOut.Length > 6) Then
                    rslt += "," + cOut(i) : i += 2 'max velocity.
                    rslt += "," + cOut(2)  'max flow.
                    rslt += "," + cOut(i) : i -= 1 'max/full depth
                    rslt += "," + cOut(i)  'max/full flow
                End If
                rptLines.Add(uID, rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'return linkID with maximum velocity, max flow, and max depth.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    Public Function SplitSpace(input As String) As String()
        Dim uList As New List(Of String)
        Dim uPattern As String = "[\s+\t]"  'split either by space Or tab.

        For Each curr As String In Regex.Split(input, uPattern)
            Dim uCurr As String = curr.Trim
            If (uCurr <> String.Empty) Then uList.Add(uCurr)
        Next

        Return uList.ToArray()
    End Function

    '  *******************
    'Node Inflow Summary
    '*******************

    '-------------------------------------------------------------------------------------------------
    '                                Maximum  Maximum                  Lateral       Total        Flow
    '                                Lateral    Total  Time Of Max      Inflow      Inflow     Balance
    '                                 Inflow   Inflow   Occurrence      Volume      Volume       Error
    'Node                 Type           MGD      MGD  days hr:min    10^6 gal    10^6 gal     Percent
    '-------------------------------------------------------------------------------------------------
    Public Function read_NodeFlowSummary(rptFile As String) As List(Of String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New List(Of String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Node Inflow Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 8
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                i = 3
                Dim rslt As String = cOut(0)
                If (cOut.Length > 2) Then
                    rslt += "," + cOut(i)  'Maximum inflow.
                End If
                rptLines.Add(rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'return the manhole number with maximum flow data.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '  ******************
    'Node Depth Summary
    '******************

    '---------------------------------------------------------------------------------
    '                               Average  Maximum  Maximum  Time Of Max    Reported
    '                                 Depth    Depth      HGL   Occurrence   Max Depth
    'Node                 Type         Feet     Feet     Feet  days hr:min        Feet
    '---------------------------------------------------------------------------------
    Public Function read_NodeDepthSummary(rptFile As String, Optional blnNodeType As Boolean = False) As Dictionary(Of String, String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New Dictionary(Of String, String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Node Depth Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 7
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                i = 3
                Dim rslt As String = cOut(0) 'Node ID.
                If (cOut.Length > 3) Then
                    rslt += "," + cOut(i) : i += 1 'max depth.
                    rslt += "," + cOut(i) 'HGL.
                    If blnNodeType Then rslt += "," + cOut(1) 'Node type
                End If
                rptLines.Add(cOut(0), rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'returns manhole number with max. depth and HGL.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '  ***********************
    'Outfall Loading Summary
    '***********************

    '-----------------------------------------------------------
    '                       Flow       Avg       Max       Total
    '                       Freq      Flow      Flow      Volume
    'Outfall Node           Pcnt       MGD       MGD    10^6 gal
    '-----------------------------------------------------------
    Public Function read_Outfall_Summary(rptFile As String) As List(Of String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New List(Of String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Outfall Loading Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 7
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                i = 3
                Dim rslt As String = cOut(0)
                If (cOut.Length > 3) Then
                    rslt += "," + cOut(i) : i += 1 'max flow.
                    rslt += "," + cOut(i) 'total volume.
                End If
                rptLines.Add(rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'returns manhole number with max. depth and total volume.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '   **********************
    'Node Surcharge Summary
    '**********************

    'Surcharging occurs When water rises above the top Of the highest conduit.
    '---------------------------------------------------------------------
    '                                             Max. Height   Min. Depth
    '                                 Hours       Above Crown    Below Rim
    'Node                 Type      Surcharged           Feet         Feet
    '---------------------------------------------------------------------
    Public Function read_NodeSurcharge_Summary(rptFile As String) As List(Of String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New List(Of String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Node Surcharge Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 8
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                i = 3
                Dim rslt As String = cOut(0)
                If (cOut.Length > 3) Then
                    rslt += "," + cOut(i) : i += 1 'max height above crown.
                    rslt += "," + cOut(i) 'min Depth below rim feet.
                End If
                rptLines.Add(rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'returns manhole number with max. height above crown and min depth below rim.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '  *********************
    'Node Flooding Summary
    '*********************

    'Flooding refers To all water that overflows a node, whether it ponds Or Not.
    '--------------------------------------------------------------------------
    '                                                           Total   Maximum
    '                               Maximum   Time Of Max       Flood    Ponded
    '                      Hours       Rate    Occurrence      Volume     Depth
    'Node                 Flooded       MGD   days hr:min    10^6 gal      Feet
    '--------------------------------------------------------------------------
    Public Function read_NodeFlooding_Summary(rptFile As String) As List(Of String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New List(Of String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Node Flooding Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 9
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
            If (txtLine = "No nodes were flooded.") Then GoTo returnNothing
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                Dim rslt As String = cOut(0)
                If (cOut.Length > 1) Then rslt += "," + cOut(1)  'Hours Flooded.
                If (cOut.Length > 5) Then rslt += "," + cOut(5) 'Total volume.
                rptLines.Add(rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'returns manhole number with Hours flooded and volume of flooding respectively.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '  ********************
    'Link Summary
    '************
    'Name             From Node        To Node          Type            Length    %Slope Roughness
    '---------------------------------------------------------------------------------------------
    Public Function read_LinkSlopes(rptFile As String) As Dictionary(Of String, String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New Dictionary(Of String, String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Link Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 3
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                Dim uID As String = cOut(0).Trim()
                Dim rslt As String = String.Empty
                If (cOut.Length > 5) Then rslt = cOut(5) '%Slope
                rptLines.Add(uID, rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'return linkID with maximum velocity, max flow, and max depth.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    '  ********************
    'Cross Section Summary
    '*********************
    '                                      Full     Full     Hyd.     Max.   No. Of     Full
    'Conduit          Shape               Depth     Area     Rad.    Width  Barrels     Flow
    '---------------------------------------------------------------------------------------
    Public Function read_LinkDiameters(rptFile As String) As Dictionary(Of String, String)
        Dim i As Integer, txtLine As String = String.Empty
        If Not File.Exists(rptFile) Then Return Nothing

        Dim rptLines As New Dictionary(Of String, String)
        Dim rptIn As StreamReader = New StreamReader(rptFile)

        Do While (Not rptIn.EndOfStream)
            txtLine = rptIn.ReadLine().Trim()
            If txtLine = "Cross Section Summary" Then GoTo exitLoop
        Loop
exitLoop:
        If (rptIn.EndOfStream) Then GoTo returnNothing
        For i = 1 To 4
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        Next i
        If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()

        While (txtLine <> String.Empty)
            Dim cOut As String() = SplitSpace(txtLine)
            If (cOut.Length > 0) Then
                Dim uID As String = cOut(0).Trim()
                Dim rslt As String = String.Empty
                If (cOut.Length > 2) Then rslt = cOut(2) 'conduit dia in ft.
                rptLines.Add(uID, rslt)
            End If
            If Not rptIn.EndOfStream Then txtLine = rptIn.ReadLine().Trim()
        End While
        rptIn.Close()
        'return linkID with maximum velocity, max flow, and max depth.
        Return rptLines
returnNothing:
        rptIn.Close()
        Return Nothing
    End Function

    Public Sub closeSWMMexecution(Optional swmmSuc As Integer = 0,
                                  Optional rptFile As String = "tmpSWMM.rpt")
        Try
            If (swmmSuc > 299) Then 'If it not the INP file parsing error.
                swmm_end()
                swmm_close()
            End If
            CloseSwmmOutFile()
            If (swmmSuc > 299) Then
                File.Delete(rptFile) 'Kill(rptFile)
            End If
        Catch
            'do nothing
        End Try
    End Sub
End Module
