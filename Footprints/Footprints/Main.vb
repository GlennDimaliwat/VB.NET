Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection.Assembly
Imports System.Reflection

<Assembly: AssemblyVersion("1.4.6.5")> 

Public Class Main

    Private buildVersion As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
    Private bindingSource As New BindingSource()
    Private selectedTicket As String
    Private selectedWorkspace As String
    Private currentSearchIndex As Integer = 0
    Private currentFindString As String
    Private selectedTabText As String
    Private selectedQueryType As String = "1"
    Private dateModifiedFromFilter As String
    Private dateModifiedToFilter As String
    Private ticketStatusAssigned As Boolean = True
    Private ticketStatusAccepted As Boolean = True
    Private ticketStatusPending As Boolean = True
    Private ticketStatusResolved As Boolean = False
    Private ticketStatusClosed As Boolean = False
    Private ticketStatusCancelled As Boolean = False
    Private ticketTypeRFC As Boolean = True
    Private ticketTypeSD As Boolean = True
    Private ticketTypePM As Boolean = True
    Private changePhaseFilterEvaluation As Boolean = True
    Private changePhaseFilterEndorsementToCab As Boolean = True
    Private changePhaseFilterBuildDev As Boolean = True
    Private changePhaseFilterUAT As Boolean = True
    Private changePhaseFilterProdTransportApproval As Boolean = True
    Private changePhaseFilterImplementation As Boolean = True

#Region "LOAD FORM"
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Load File Path
        filepathTextBox.Text = My.Settings.FilePath

        'Load Username
        If Environment.UserName = "INDRA" And Environment.UserDomainName = "888-00256" Then
            UsernameToolStripTextBox.Text = "GDIMALIWAT"
        Else
            UsernameToolStripTextBox.Text = Environment.UserName
        End If

        'Load Status Bar
        StatusBarText.Text = "Your machine is currently logged in as " + Environment.UserName + " on domain " + Environment.UserDomainName
    End Sub
#End Region

#Region "REFRESH"
    Private Sub RefreshAction(ByVal queryType As String)

        Dim dataList As New SortableBindingList(Of Data)

        Try
            Dim connToMySQL As New ConnectionHandler
            Dim stringQuery As String = ""
            Dim data As DataTable = Nothing
            Dim conditionMaster1 As String
            Dim conditionMaster2 As String
            Dim conditionMaster3 As String
            Dim fs As New FixStrings()
            Dim fm As New FileManager()

            If QueryTypeToolStripComboBox.SelectedItem Is Nothing Then
                MsgBox("Please select a Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
                Exit Sub
            End If

            If queryType = "1" Then
                If File.Exists(filepathTextBox.Text) = True Then
                    Dim inputTicketList As List(Of String) = fm.readFile(filepathTextBox.Text)

                    If (inputTicketList.Count > 0) Then
                        conditionMaster1 = ""
                        conditionMaster2 = ""
                        conditionMaster3 = ""

                        For i As Integer = 0 To inputTicketList.Count - 1
                            If inputTicketList(i).Contains("SD") And ticketTypeSD = True Then
                                If conditionMaster1.Length <> 0 Then
                                    conditionMaster1 = conditionMaster1 & ","
                                End If
                                conditionMaster1 = conditionMaster1 & "'" & inputTicketList(i).Replace("SD", "") & "'"

                            ElseIf inputTicketList(i).Contains("PM") And ticketTypePM = True Then
                                If conditionMaster2.Length <> 0 Then
                                    conditionMaster2 = conditionMaster2 & ","
                                End If
                                conditionMaster2 = conditionMaster2 & "'" & inputTicketList(i).Replace("PM", "") & "'"

                            ElseIf inputTicketList(i).Contains("RFC") And ticketTypeRFC = True Then
                                If conditionMaster3.Length <> 0 Then
                                    conditionMaster3 = conditionMaster3 & ","
                                End If
                                conditionMaster3 = conditionMaster3 & "'" & inputTicketList(i).Replace("RFC", "") & "'"
                            End If
                        Next

                        connToMySQL.connect()

                        stringQuery = ""
                        If conditionMaster1 <> "" Then
                            stringQuery = "SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'SD '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, NULL, NULL, NULL, NULL, NULL, NULL"
                            stringQuery &= " FROM MASTER1 a"
                            stringQuery &= " INNER JOIN MASTER1_ABDATA b ON a.mrID = b.mrID"
                            stringQuery &= " WHERE a.mrID IN (" & conditionMaster1 & ")"
                            If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                                stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                            End If
                        End If

                        If conditionMaster2 <> "" Then
                            If conditionMaster1 <> "" Then
                                stringQuery &= " UNION ALL"
                            End If
                            stringQuery &= " SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'PM '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, NULL, NULL, NULL, NULL, NULL, NULL"
                            stringQuery &= " FROM MASTER2 a"
                            stringQuery &= " INNER JOIN MASTER2_ABDATA b ON a.mrID = b.mrID"
                            stringQuery &= " WHERE a.mrID IN (" & conditionMaster2 & ")"
                            If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                                stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                            End If
                        End If

                        If conditionMaster3 <> "" Then
                            If conditionMaster1 <> "" Or conditionMaster2 <> "" Then
                                stringQuery &= " UNION ALL"
                            End If
                            stringQuery &= " SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'RFC '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, Change__bPhase, (CONVERT(VARCHAR(10), a.Review__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Review__bDate, 108)) as REVIEW_DATE, (CONVERT(VARCHAR(10), a.Authorization__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Authorization__bDate, 108)) as AUTHORIZATION_DATE, (CONVERT(VARCHAR(10), a.Build__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Build__bDate, 108)) as BUILD_DATE, (CONVERT(VARCHAR(10), a.Testing__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Testing__bDate, 108)) as TESTING_DATE, (CONVERT(VARCHAR(10), a.Implementation__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Implementation__bDate, 108)) as IMPLEM_DATE"
                            stringQuery &= " FROM MASTER3 a"
                            stringQuery &= " INNER JOIN MASTER3_ABDATA b ON a.mrID = b.mrID"
                            stringQuery &= " WHERE a.mrID IN (" & conditionMaster3 & ")"
                            If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                                stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                            End If
                        End If
                        If stringQuery <> "" Then
                            stringQuery &= " ORDER BY a.Ticket__bSummary ASC"
                        End If

                        If stringQuery <> "" Then
                            data = connToMySQL.query(stringQuery)
                        End If
                        connToMySQL.close()
                        connToMySQL = Nothing

                    Else
                        MsgBox("The tickets file is empty", MsgBoxStyle.OkOnly, "Footprints Viewer")
                        Exit Sub
                    End If
                Else
                    MsgBox("Unable to read file " & filepathTextBox.Text, MsgBoxStyle.OkOnly, "Footprints Viewer")
                    Exit Sub
                End If

            ElseIf queryType = "2" Then
                If SubcategoryToolStripComboBox.Text <> "" Then
                    connToMySQL.connect()

                    stringQuery = ""

                    If ticketTypeSD = True Then
                        stringQuery = "SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'SD '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, NULL, NULL, NULL, NULL, NULL, NULL"
                        stringQuery &= " FROM MASTER1 a"
                        stringQuery &= " INNER JOIN MASTER1_ABDATA b ON a.mrID = b.mrID"
                        stringQuery &= " WHERE LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end) LIKE '%" + fs.unfixString(UsernameToolStripTextBox.Text.Trim) + "%'"
                        If SubcategoryToolStripComboBox.Text <> "All" Then
                            stringQuery &= " AND a.Sub__bCategory = '" + fs.unfixString(SubcategoryToolStripComboBox.Text.Trim) + "'"
                        End If
                        If DescriptionToolStripTextBox.Text.Trim <> "" Then
                            stringQuery &= " AND a.mrTitle LIKE '%" + fs.unfixString(DescriptionToolStripTextBox.Text.Trim) + "%'"
                        End If
                        If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                            stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                        End If
                        stringQuery &= " AND ( a.mrStatus = ''"
                        If ticketStatusAssigned = True Or ticketStatusAccepted = True Or ticketStatusPending = True Or ticketStatusResolved = True Or ticketStatusClosed = True Or ticketStatusCancelled = True Then
                            If ticketStatusAssigned = True Then
                                stringQuery &= " or a.mrStatus = 'Assigned'"
                            End If
                            If ticketStatusAccepted = True Then
                                stringQuery &= " or a.mrStatus = 'Accepted'"
                            End If
                            If ticketStatusPending = True Then
                                stringQuery &= " or a.mrStatus = 'Pending'"
                            End If
                            If ticketStatusResolved = True Then
                                stringQuery &= " or a.mrStatus = 'Resolved'"
                            End If
                            If ticketStatusClosed = True Then
                                stringQuery &= " or a.mrStatus = 'Closed'"
                            End If
                            If ticketStatusCancelled = True Then
                                stringQuery &= " or a.mrStatus = 'Cancelled'"
                            End If
                        End If
                        stringQuery &= " )"
                    End If

                    If stringQuery <> "" And ticketTypeSD = True And (ticketTypePM = True Or ticketTypeRFC = True) Then
                        stringQuery &= " UNION ALL"
                    End If

                    If ticketTypePM = True Then
                        stringQuery &= " SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'PM '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, NULL, NULL, NULL, NULL, NULL, NULL"
                        stringQuery &= " FROM MASTER2 a"
                        stringQuery &= " INNER JOIN MASTER2_ABDATA b ON a.mrID = b.mrID"
                        stringQuery &= " WHERE LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end) LIKE '%" + fs.unfixString(UsernameToolStripTextBox.Text.Trim) + "%'"
                        If SubcategoryToolStripComboBox.Text <> "All" Then
                            stringQuery &= " AND a.Sub__bCategory = '" + fs.unfixString(SubcategoryToolStripComboBox.Text.Trim) + "'"
                        End If
                        If DescriptionToolStripTextBox.Text.Trim <> "" Then
                            stringQuery &= " AND a.mrTitle LIKE '%" + fs.unfixString(DescriptionToolStripTextBox.Text.Trim) + "%'"
                        End If
                        If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                            stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                        End If
                        stringQuery &= " AND ( a.mrStatus = ''"
                        If ticketStatusAssigned = True Or ticketStatusAccepted = True Or ticketStatusPending = True Or ticketStatusResolved = True Or ticketStatusClosed = True Or ticketStatusCancelled = True Then
                            If ticketStatusAssigned = True Then
                                stringQuery &= " or a.mrStatus = 'Assigned'"
                            End If
                            If ticketStatusAccepted = True Then
                                stringQuery &= " or a.mrStatus = 'Accepted'"
                            End If
                            If ticketStatusPending = True Then
                                stringQuery &= " or a.mrStatus = 'Pending'"
                            End If
                            If ticketStatusResolved = True Then
                                stringQuery &= " or a.mrStatus = 'Resolved'"
                            End If
                            If ticketStatusClosed = True Then
                                stringQuery &= " or a.mrStatus = 'Closed'"
                            End If
                            If ticketStatusCancelled = True Then
                                stringQuery &= " or a.mrStatus = 'Cancelled'"
                            End If
                        End If
                        stringQuery &= " )"
                    End If

                    If stringQuery <> "" And stringQuery.EndsWith("UNION ALL") = False And (ticketTypeSD = True Or ticketTypePM = True) And ticketTypeRFC = True Then
                        stringQuery &= " UNION ALL"
                    End If

                    If ticketTypeRFC = True Then
                        stringQuery &= " SELECT a.Ticket__bSummary, (CONVERT(VARCHAR(10), a.mrSUBMITDATE, 101) + ' ' + CONVERT(VARCHAR(8), a.mrSUBMITDATE, 108)) as DATE_CREATED, (CONVERT(VARCHAR(10), a.SLA__bDue__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.SLA__bDue__bDate, 108)) AS DUE_DATE, a.Category, a.Sub__bCategory, b.Display__bName, 'RFC '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end), (CONVERT(VARCHAR(10), a.mrUpdateDate, 101) + ' ' + CONVERT(VARCHAR(8), a.mrUpdateDate, 108)) as UPDATE_DATE, a.mrALLDESCRIPTIONS, a.T__ucode, Change__bPhase, (CONVERT(VARCHAR(10), a.Review__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Review__bDate, 108)) as REVIEW_DATE, (CONVERT(VARCHAR(10), a.Authorization__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Authorization__bDate, 108)) as AUTHORIZATION_DATE, (CONVERT(VARCHAR(10), a.Build__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Build__bDate, 108)) as BUILD_DATE, (CONVERT(VARCHAR(10), a.Testing__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Testing__bDate, 108)) as TESTING_DATE, (CONVERT(VARCHAR(10), a.Implementation__bDate, 101) + ' ' + CONVERT(VARCHAR(8), a.Implementation__bDate, 108)) as IMPLEM_DATE"
                        stringQuery &= " FROM MASTER3 a"
                        stringQuery &= " INNER JOIN MASTER3_ABDATA b ON a.mrID = b.mrID"
                        stringQuery &= " WHERE LEFT(a.mrAssignees, case when CHARINDEX('CC:', a.mrAssignees ) = 0 then LEN(a.mrAssignees) else CHARINDEX('CC:', a.mrAssignees) -1 end) LIKE '%" + fs.unfixString(UsernameToolStripTextBox.Text.Trim) + "%'"
                        If SubcategoryToolStripComboBox.Text <> "All" Then
                            stringQuery &= " AND a.Sub__bCategory = '" + fs.unfixString(SubcategoryToolStripComboBox.Text.Trim) + "'"
                        End If
                        If DescriptionToolStripTextBox.Text.Trim <> "" Then
                            stringQuery &= " AND a.mrTitle LIKE '%" + fs.unfixString(DescriptionToolStripTextBox.Text.Trim) + "%'"
                        End If
                        If (SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked) Then
                            stringQuery &= " AND ( a.mrUpdateDate >= '" + dateModifiedFromFilter + "' AND a.mrUpdateDate <= '" + dateModifiedToFilter + "')"
                        End If
                        stringQuery &= " AND ( a.mrStatus = ''"
                        If ticketStatusAssigned = True Or ticketStatusAccepted = True Or ticketStatusPending = True Or ticketStatusResolved = True Or ticketStatusClosed = True Or ticketStatusCancelled = True Then
                            If ticketStatusAssigned = True Then
                                stringQuery &= " or a.mrStatus = 'Assigned'"
                            End If
                            If ticketStatusAccepted = True Then
                                stringQuery &= " or a.mrStatus = 'Accepted'"
                            End If
                            If ticketStatusPending = True Then
                                stringQuery &= " or a.mrStatus = 'Pending'"
                            End If
                            If ticketStatusResolved = True Then
                                stringQuery &= " or a.mrStatus = 'Resolved'"
                            End If
                            If ticketStatusClosed = True Then
                                stringQuery &= " or a.mrStatus = 'Closed'"
                            End If
                            If ticketStatusCancelled = True Then
                                stringQuery &= " or a.mrStatus = 'Cancelled'"
                            End If
                        End If
                        stringQuery &= " )"
                    End If

                    If stringQuery <> "" Then
                        stringQuery &= " ORDER BY a.Ticket__bSummary ASC"
                    End If

                    'System.Diagnostics.Debug.WriteLine(stringQuery)

                    If stringQuery <> "" Then
                        data = connToMySQL.query(stringQuery)
                    End If

                    connToMySQL.close()
                    connToMySQL = Nothing
                Else
                    MsgBox("Please select a Sub Category", MsgBoxStyle.OkOnly, "Footprints Viewer")
                    Exit Sub
                End If

            End If

            'MsgBox(data.Rows.Count)
            If stringQuery <> "" Then
                If data.Rows.Count > 0 Then
                    For i As Integer = 0 To data.Rows.Count - 1
                        Dim ticketNumber As String = fs.fixString(data.Rows(i)(6).ToString)
                        Dim currentWorkspace As String
                        If ticketNumber.StartsWith("SD") Then
                            currentWorkspace = "1"
                        ElseIf ticketNumber.StartsWith("PM") Then
                            currentWorkspace = "2"
                        ElseIf ticketNumber.StartsWith("RFC") Then
                            currentWorkspace = "3"
                        Else
                            currentWorkspace = "1"
                        End If

                        Dim progress As String = ""
                        If File.Exists("Progress.dat") Then
                            Dim progressFileLines() As String
                            progressFileLines = File.ReadAllLines("Progress.dat")
                            For Each line As String In progressFileLines
                                Dim lineTicketNumber As String = line.Split(" ")(0) + " " + line.Split(" ")(1)
                                'If line.Contains(ticketNumber) = True Then
                                If lineTicketNumber = ticketNumber Then
                                    'System.Diagnostics.Debug.WriteLine(lineTicketNumber + " vs " + ticketNumber)
                                    progress = line.Split(" ")(2)
                                    Exit For
                                End If
                            Next
                        End If

                        Dim ticketStatus As String = fs.fixString(data.Rows(i)(8).ToString)
                        Dim remarks As String

                        If ticketStatus <> "Closed" Then
                            If fs.fixString(data.Rows(i)(9).ToString) <> "" Then
                                remarks = fs.fixString(data.Rows(i)(9).ToString)
                            Else
                                remarks = fs.fixString(data.Rows(i)(10).ToString)
                            End If
                        Else
                            remarks = ""
                        End If

                        Dim dateResolved As String = ""
                        If ticketStatus = "Resolved" Or ticketStatus = "Closed" Then
                            dateResolved = getDateResolved(ticketNumber, currentWorkspace, "Resolved")
                            If dateResolved = "" Then
                                dateResolved = getDateResolved(ticketNumber, currentWorkspace, "Closed")
                            End If
                        End If

                        Dim kbSolutionNumber As String = ""
                        Dim dateLinkedToSolution As String = ""
                        Dim comments As String = fs.fixString(data.Rows(i)(14).ToString)
                        getKB(ticketNumber, currentWorkspace, comments, kbSolutionNumber, dateLinkedToSolution)

                        Dim tCode As String = fs.fixString(data.Rows(i)(15).ToString)
                        Dim changePhase As String = fs.fixString(data.Rows(i)(16).ToString)

                        Dim lastUpdated As String = getTimeElapsed(data.Rows(i)(13).ToString)
                        'System.Diagnostics.Debug.WriteLine(lastUpdated)

                        'Filter Out Default Value
                        Dim filterOut = False

                        If changePhase = "Evaluation" Then
                            If changePhaseFilterEvaluation = False Then
                                filterOut = True
                            End If
                        ElseIf changePhase = "Endorsement to CAB" Then
                            If changePhaseFilterEndorsementToCab = False Then
                                filterOut = True
                            End If
                        ElseIf changePhase = "Build/Dev" Then
                            If changePhaseFilterBuildDev = False Then
                                filterOut = True
                            End If
                        ElseIf changePhase = "UAT" Then
                            If changePhaseFilterUAT = False Then
                                filterOut = True
                            End If
                        ElseIf changePhase = "Prod Transport Approval" Then
                            If changePhaseFilterProdTransportApproval = False Then
                                filterOut = True
                            End If
                        ElseIf changePhase = "Implementation" Then
                            If changePhaseFilterImplementation = False Then
                                filterOut = True
                            End If
                        End If

                        If filterOut = False Then
                            dataList.Add(New Data(fs.fixString(data.Rows(i)(0).ToString),
                                                    progress, fs.fixString(data.Rows(i)(1).ToString),
                                                    fs.fixString(data.Rows(i)(2).ToString),
                                                    fs.fixString(data.Rows(i)(3).ToString),
                                                    fs.fixString(data.Rows(i)(4).ToString),
                                                    fs.fixString(data.Rows(i)(5).ToString),
                                                    ticketNumber,
                                                    fs.fixString(data.Rows(i)(7).ToString),
                                                    ticketStatus,
                                                    remarks,
                                                    fs.fixString(data.Rows(i)(11).ToString),
                                                    fs.fixString(data.Rows(i)(12).ToString),
                                                    fs.fixString(data.Rows(i)(13).ToString),
                                                    dateResolved,
                                                    kbSolutionNumber,
                                                    dateLinkedToSolution,
                                                    fs.fixString(tCode),
                                                    fs.fixString(changePhase),
                                                    fs.fixString(data.Rows(i)(17).ToString),
                                                    fs.fixString(data.Rows(i)(18).ToString),
                                                    fs.fixString(data.Rows(i)(19).ToString),
                                                    fs.fixString(data.Rows(i)(20).ToString),
                                                    fs.fixString(data.Rows(i)(21).ToString),
                                                    lastUpdated))
                        End If
                    Next
                End If

                ' Populate a new data table and bind it to the BindingSource.
                Me.DataGridView1.DataSource = dataList
                colorGrids()

                ' Center PROGRESS text
                Me.DataGridView1.Columns("PROGRESS").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                ' Adjust Header Labels and Tooltips
                Me.DataGridView1.Columns("TICKET_SUMMARY").HeaderText = "TICKET SUMMARY"
                Me.DataGridView1.Columns("TICKET_SUMMARY").ToolTipText = "TICKET SUMMARY"
                Me.DataGridView1.Columns("PROGRESS").HeaderText = "PROGRESS"
                Me.DataGridView1.Columns("PROGRESS").ToolTipText = "PROGRESS"
                Me.DataGridView1.Columns("CREATED").HeaderText = "DATE CREATED"
                Me.DataGridView1.Columns("CREATED").ToolTipText = "DATE CREATED"
                Me.DataGridView1.Columns("DUE_DATE").HeaderText = "DUE DATE"
                Me.DataGridView1.Columns("DUE_DATE").ToolTipText = "DUE DATE"
                Me.DataGridView1.Columns("CATEGORY").HeaderText = "CATEGORY"
                Me.DataGridView1.Columns("CATEGORY").ToolTipText = "CATEGORY"
                Me.DataGridView1.Columns("SUB_CATEGORY").HeaderText = "SUB CATEGORY"
                Me.DataGridView1.Columns("SUB_CATEGORY").ToolTipText = "SUB CATEGORY"
                Me.DataGridView1.Columns("REQUESTOR").HeaderText = "REQUESTOR"
                Me.DataGridView1.Columns("REQUESTOR").ToolTipText = "REQUESTOR"
                Me.DataGridView1.Columns("TICKET").HeaderText = "TICKET NUMBER"
                Me.DataGridView1.Columns("TICKET").ToolTipText = "TICKET NUMBER"
                Me.DataGridView1.Columns("DESCRIPTION").HeaderText = "DESCRIPTION"
                Me.DataGridView1.Columns("DESCRIPTION").ToolTipText = "DESCRIPTION"
                Me.DataGridView1.Columns("STATUS").HeaderText = "STATUS"
                Me.DataGridView1.Columns("STATUS").ToolTipText = "STATUS"
                Me.DataGridView1.Columns("REMARKS").HeaderText = "REMARKS"
                Me.DataGridView1.Columns("REMARKS").ToolTipText = "REMARKS"
                Me.DataGridView1.Columns("CLASSIFICATION").HeaderText = "CLASSIFICATION"
                Me.DataGridView1.Columns("CLASSIFICATION").ToolTipText = "CLASSIFICATION"
                Me.DataGridView1.Columns("ASSIGNEES").HeaderText = "ASSIGNEES"
                Me.DataGridView1.Columns("ASSIGNEES").ToolTipText = "ASSIGNEES"
                Me.DataGridView1.Columns("MODIFIED").HeaderText = "DATE MODIFIED"
                Me.DataGridView1.Columns("MODIFIED").ToolTipText = "DATE MODIFIED"
                Me.DataGridView1.Columns("RESOLVED").HeaderText = "DATE RESOLVED"
                Me.DataGridView1.Columns("RESOLVED").ToolTipText = "DATE RESOLVED"
                Me.DataGridView1.Columns("KB").HeaderText = "KNOWLEDGE BASE"
                Me.DataGridView1.Columns("KB").ToolTipText = "KNOWLEDGE BASE"
                Me.DataGridView1.Columns("DATE_LINKED_TO_KB").HeaderText = "DATE LINKED TO KB"
                Me.DataGridView1.Columns("DATE_LINKED_TO_KB").ToolTipText = "DATE LINKED TO KB"
                Me.DataGridView1.Columns("T_CODE").HeaderText = "TRANSACTION CODE"
                Me.DataGridView1.Columns("T_CODE").ToolTipText = "TRANSACTION CODE"
                Me.DataGridView1.Columns("CHANGE_PHASE").HeaderText = "CHANGE PHASE"
                Me.DataGridView1.Columns("CHANGE_PHASE").ToolTipText = "CHANGE PHASE"
                Me.DataGridView1.Columns("REVIEW_DATE").HeaderText = "REVIEW DATE"
                Me.DataGridView1.Columns("REVIEW_DATE").ToolTipText = "REVIEW DATE"
                Me.DataGridView1.Columns("AUTHORIZATION_DATE").HeaderText = "AUTHORIZATION DATE"
                Me.DataGridView1.Columns("AUTHORIZATION_DATE").ToolTipText = "AUTHORIZATION DATE"
                Me.DataGridView1.Columns("BUILD_DATE").HeaderText = "BUILD DATE"
                Me.DataGridView1.Columns("BUILD_DATE").ToolTipText = "BUILD DATE"
                Me.DataGridView1.Columns("TESTING_DATE").HeaderText = "TESTING DATE"
                Me.DataGridView1.Columns("TESTING_DATE").ToolTipText = "TESTING DATE"
                Me.DataGridView1.Columns("IMPLEM_DATE").HeaderText = "IMPLEMENTATION DATE"
                Me.DataGridView1.Columns("IMPLEM_DATE").ToolTipText = "IMPLEMENTATION DATE"
                Me.DataGridView1.Columns("LAST_UPDATED").HeaderText = "LAST UPDATED"
                Me.DataGridView1.Columns("LAST_UPDATED").ToolTipText = "LAST UPDATED"

                'Toggle Columns
                ToggleColumns()

                StatusBarText.Text = data.Rows.Count.ToString + " tickets extracted"
            Else
                MsgBox("The query does not contain any valid tickets", MsgBoxStyle.OkOnly, "Footprints Viewer")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "Footprints Viewer")
        Finally
            dataList = Nothing
        End Try
    End Sub
#End Region

#Region "GET KB"
    Sub getKB(ByVal ticketNumber As String, ByVal currentWorkspace As String, ByVal comments As String, ByRef kbSolutionNumber As String, ByRef dateLinkedToSolution As String)
        Dim connToMySQL As New ConnectionHandler
        Dim dataHistory As DataTable
        Dim stringQuery As String = ""
        Dim kbFound As Boolean = False
        Dim rfcFound As Boolean = False
        Dim docsFound As Boolean = False

        connToMySQL.connect()

        stringQuery = "SELECT mrHISTORY"
        stringQuery &= " FROM MASTER" + currentWorkspace + "_HISTORY"
        stringQuery &= " WHERE mrID = '" + ticketNumber.Split(" ")(1) + "'"
        stringQuery &= " AND mrHistory like '%Copied or Selected from add%'"

        dataHistory = connToMySQL.query(stringQuery)
        connToMySQL.close()
        connToMySQL = Nothing

        If dataHistory.Rows.Count > 0 Then
            For i As Integer = 0 To dataHistory.Rows.Count - 1
                kbSolutionNumber = dataHistory.Rows(i)(0).ToString.Split(" ")(9)
                dateLinkedToSolution = dataHistory.Rows(i)(0).ToString.Split(" ")(0) + " " + dataHistory.Rows(i)(0).ToString.Split(" ")(1)

                If IsNumeric(kbSolutionNumber) Then
                    kbSolutionNumber = "SOLUTION " + kbSolutionNumber
                    kbFound = True
                Else
                    If kbSolutionNumber.Contains("X3") Then
                        kbSolutionNumber = "RFC " + kbSolutionNumber.Replace("l", "").Replace("X3", "")
                        rfcFound = True
                    Else
                        kbSolutionNumber = ""
                        dateLinkedToSolution = ""
                    End If
                End If

                If comments.Contains("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/") Then
                    docsFound = True
                End If
                If kbFound = False And rfcFound = False And docsFound = True Then
                    kbSolutionNumber = "TA/TP"
                End If

            Next
        End If

    End Sub
#End Region

#Region "GET TIME ELAPSED"
    Function getTimeElapsed(ByVal dateModifed As String)
        Dim lastUpdated As String = ""
        Dim currentDate As DateTime = DateTime.Now
        Dim modifiedDate As DateTime = DateTime.ParseExact(dateModifed, "MM/dd/yyyy HH:mm:ss", Nothing)
        Dim timeSpan As TimeSpan = currentDate.Subtract(modifiedDate)
        Dim weekendDays As Integer = getWeekendDays(currentDate.Date, modifiedDate.Date)
        Dim actualDays As Integer = timeSpan.Days
        Dim actualHours As Integer = timeSpan.Hours
        Dim actualMinutes As Integer = timeSpan.Minutes

        If weekendDays > 0 And actualDays > 0 Then
            actualDays = actualDays - weekendDays
        End If

        If actualDays = 0 Or actualDays = 1 Then
            lastUpdated = actualDays.ToString + " Day"
            If actualHours = 0 Or actualHours = 1 Then
                lastUpdated += ", " + actualHours.ToString + " Hour"
            ElseIf actualHours > 1 Then
                lastUpdated += ", " + actualHours.ToString + " Hours"
            End If
            If actualMinutes = 0 Or actualMinutes = 1 Then
                lastUpdated += ", " + actualMinutes.ToString + " Minute"
            ElseIf actualMinutes > 1 Then
                lastUpdated += ", " + actualMinutes.ToString + " Minutes"
            End If
        ElseIf actualDays > 1 Then
            lastUpdated = actualDays.ToString + " Days"
            If actualHours = 0 Or actualHours = 1 Then
                lastUpdated += ", " + actualHours.ToString + " Hour"
            ElseIf actualHours > 1 Then
                lastUpdated += ", " + actualHours.ToString + " Hours"
            End If
            If actualMinutes = 0 Or actualMinutes = 1 Then
                lastUpdated += ", " + actualMinutes.ToString + " Minute"
            ElseIf actualMinutes > 1 Then
                lastUpdated += ", " + actualMinutes.ToString + " Minutes"
            End If
        End If
        lastUpdated += " ago"
        'lastUpdated += weekendDays.ToString + "weekends days"

        Return lastUpdated
    End Function

    Function getWeekendDays(ByRef startDate As Date, ByRef endDate As Date) As Integer
        Dim workingStartDate As Date = startDate
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        WeekendDays = 0

        totalDays = DateDiff(DateInterval.Day, workingStartDate, endDate) * -1
        'System.Diagnostics.Debug.WriteLine("workingStartDate: " + workingStartDate.ToString)
        'System.Diagnostics.Debug.WriteLine("endDate: " + endDate.ToString)
        'System.Diagnostics.Debug.WriteLine("Total Days: " + totalDays.ToString)

        For i As Integer = 1 To totalDays
            'System.Diagnostics.Debug.WriteLine("Loop workingStartDate: " + workingStartDate.ToString)
            If DatePart(DateInterval.Weekday, workingStartDate) = 1 Then 'Sunday
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, workingStartDate) = 7 Then 'Saturday
                WeekendDays = WeekendDays + 1
            End If
            workingStartDate = DateAdd("d", -1, workingStartDate)
        Next
        'System.Diagnostics.Debug.WriteLine("WeekendDays: " + WeekendDays.ToString)
        Return WeekendDays
    End Function
#End Region

#Region "GET DATE RESOLVED"
    Function getDateResolved(ByVal ticketNumber As String, ByVal currentWorkspace As String, ByVal ticketStatus As String) As String
        Dim connToMySQL As New ConnectionHandler
        Dim dataHistory As DataTable
        Dim dateResolved As String
        Dim stringQuery As String = ""
        Dim mrGeneration As Integer = 0

        connToMySQL.connect()

        stringQuery = "SELECT mrHISTORY, mrGeneration FROM MASTER" + currentWorkspace + "_HISTORY "
        stringQuery &= " WHERE mrID = '" + ticketNumber.Split(" ")(1) + "'"
        stringQuery &= " AND mrHistory like '%" + ticketStatus + "%'"
        stringQuery &= " ORDER BY mrGENERATION desc"

        dataHistory = connToMySQL.query(stringQuery)
        connToMySQL.close()
        connToMySQL = Nothing

        If dataHistory.Rows.Count > 0 Then
            For i As Integer = 0 To dataHistory.Rows.Count - 1
                mrGeneration = dataHistory.Rows(i)(1).ToString
                dateResolved = dataHistory.Rows(i)(0).ToString

                If mrGeneration = 0 Then
                    mrGeneration = dataHistory.Rows(i)(1).ToString
                    dateResolved = dataHistory.Rows(i)(0).ToString
                Else
                    If mrGeneration - 1 = dataHistory.Rows(i)(1).ToString Then
                        mrGeneration = dataHistory.Rows(i)(1).ToString
                        dateResolved = dataHistory.Rows(i)(0).ToString
                    Else
                        Return dateResolved.Split(" ")(0) + " " + dateResolved.Split(" ")(1)
                    End If
                End If
            Next
        End If

        Return ""
    End Function
#End Region

#Region "CREATE SUB CATEGORY ITEMS"
    Sub createSubCategoryItems()
        Dim connToMySQL As New ConnectionHandler
        Dim dataSubCategories As DataTable
        Dim stringQuery As String = ""
        Dim fs As New FixStrings()

        connToMySQL.connect()

        stringQuery = "SELECT DISTINCT Sub__bCategory"
        stringQuery &= " FROM MASTER3"
        stringQuery &= " WHERE Sub__bCategory IS NOT NULL"
        'stringQuery &= " AND mrAssignees like '%" + UsernameToolStripTextBox.Text + "%'"
        stringQuery &= " AND ( mrStatus = 'Accepted' or mrStatus = 'Assigned' or mrStatus = 'Pending' ) "

        dataSubCategories = connToMySQL.query(stringQuery)
        connToMySQL.close()
        connToMySQL = Nothing

        If dataSubCategories.Rows.Count > 0 Then
            SubcategoryToolStripComboBox.Items.Clear()
            SubcategoryToolStripComboBox.Items.Add("All")
            For i As Integer = 0 To dataSubCategories.Rows.Count - 1
                ' Add Sub Category
                SubcategoryToolStripComboBox.Items.Add(fs.fixString(dataSubCategories.Rows(i)(0).ToString))
            Next
        End If

    End Sub
#End Region

#Region "Action Events"
    Private Sub menuItem_Exit()
        Close()
    End Sub

    Private Sub RefreshToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripButton.Click
        RefreshAction(selectedQueryType)
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        RefreshAction(selectedQueryType)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub RefreshToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem1.Click
        RefreshAction(selectedQueryType)
    End Sub

    Private Sub UsernameToolStripTextBox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UsernameToolStripTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            RefreshAction(selectedQueryType)
        End If
    End Sub

    Private Sub DescriptionToolStripTextBox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DescriptionToolStripTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            RefreshAction(selectedQueryType)
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Footprints Viewer (Build Version " & buildVersion & ")" & vbNewLine & vbNewLine & "Formerly known as KBOX Viewer and Service Cloud." & vbNewLine & "Not affiliated with BMC Footprints or Dell KACE KBOX." & vbNewLine & "© 2016 Glenn Dimaliwat" & vbNewLine & "All rights reserved." & vbNewLine & vbNewLine & "This software can be freely modified and distributed to an unlimited number of machines. However, this software cannot be sold or used for profit.", MsgBoxStyle.OkOnly, "Footprints Viewer")
    End Sub
#End Region

#Region "Sort Data Grid View"
    Private Sub dataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, _
    ByVal e As DataGridViewCellMouseEventArgs) _
    Handles DataGridView1.ColumnHeaderMouseClick

        Dim newColumn As DataGridViewColumn = _
            DataGridView1.Columns(e.ColumnIndex)
        Dim oldColumn As DataGridViewColumn = DataGridView1.SortedColumn
        Dim direction As ListSortDirection

        ' If oldColumn is null, then the DataGridView is not currently sorted.
        If oldColumn IsNot Nothing Then

            ' Sort the same column again, reversing the SortOrder.
            If oldColumn Is newColumn AndAlso DataGridView1.SortOrder = _
                SortOrder.Ascending Then
                direction = ListSortDirection.Descending
            Else

                ' Sort a new column and remove the old SortGlyph.
                direction = ListSortDirection.Ascending
                oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None
            End If
        Else
            direction = ListSortDirection.Ascending
        End If

        ' Sort the selected column.
        DataGridView1.Sort(newColumn, direction)
        If direction = ListSortDirection.Ascending Then
            newColumn.HeaderCell.SortGlyphDirection = SortOrder.Ascending
        Else
            newColumn.HeaderCell.SortGlyphDirection = SortOrder.Descending
        End If

        colorGrids()
        DataGridView1.ClearSelection()
        selectedTicket = ""
    End Sub

    Private Sub dataGridView1_DataBindingComplete(ByVal sender As Object, _
        ByVal e As DataGridViewBindingCompleteEventArgs) _
        Handles DataGridView1.DataBindingComplete

        ' Put each of the columns into programmatic sort mode.
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.SortMode = DataGridViewColumnSortMode.Programmatic
        Next
    End Sub
#End Region

#Region "Color Grids"
    Private Sub colorGrids()
        ' Color the Grids
        For i As Integer = 0 To DataGridView1.Rows.Count - 1

            For colNo As Integer = 9 To 10
                'DataGridView1.Rows(i).Cells(colNo).Style.BackColor = Color.Blue
                'DataGridView1.Rows(i).Cells(colNo).Style.ForeColor = Color.White

                If DataGridView1.Rows(i).Cells(3).Value <> "" Then
                    Dim dueDateString As String = DataGridView1.Rows(i).Cells(3).Value
                    Dim dueDate = DateTime.ParseExact(dueDateString, "MM/dd/yyyy HH:mm:ss", Nothing)
                    Dim tomorrowDate = DateTime.Now.Date.AddDays(1)

                    If dueDate.Date = DateTime.Now.Date Then
                        DataGridView1.Rows(i).Cells(colNo).Style.BackColor = Color.Orange
                        DataGridView1.Rows(i).Cells(colNo).Style.ForeColor = Color.Black
                        DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Orange
                        DataGridView1.Rows(i).Cells(3).Style.ForeColor = Color.Black
                    ElseIf dueDate.Date = tomorrowDate Then
                        DataGridView1.Rows(i).Cells(colNo).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(i).Cells(colNo).Style.ForeColor = Color.Black
                        DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Yellow
                        DataGridView1.Rows(i).Cells(3).Style.ForeColor = Color.Black
                    ElseIf dueDate < DateTime.Now Then
                        DataGridView1.Rows(i).Cells(colNo).Style.BackColor = Color.Red
                        DataGridView1.Rows(i).Cells(colNo).Style.ForeColor = Color.White
                        DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Red
                        DataGridView1.Rows(i).Cells(3).Style.ForeColor = Color.White
                    End If
                End If

                If DataGridView1.Rows(i).Cells(9).Value = "Resolved" Or DataGridView1.Rows(i).Cells(9).Value = "Closed" Or DataGridView1.Rows(i).Cells(9).Value = "Cancelled" Then
                    DataGridView1.Rows(i).Cells(colNo).Style.BackColor = Color.Green
                    DataGridView1.Rows(i).Cells(colNo).Style.ForeColor = Color.White
                    DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.White
                    DataGridView1.Rows(i).Cells(3).Style.ForeColor = Color.Black
                Else
                    If DataGridView1.Rows(i).Cells(11).Value = "Incident" Then
                        DataGridView1.Rows(i).Cells(11).Style.BackColor = Color.Red
                        DataGridView1.Rows(i).Cells(11).Style.ForeColor = Color.White
                    End If
                End If

                Dim reviewDate = ""
                Dim buildDate = ""
                Dim testingDate = ""
                Dim implemDate = ""

                If DataGridView1.Rows(i).Cells(19).Value <> "" Then
                    Dim reviewDateString As String = DataGridView1.Rows(i).Cells(19).Value
                    reviewDate = DateTime.ParseExact(reviewDateString, "MM/dd/yyyy HH:mm:ss", Nothing)
                End If
                If DataGridView1.Rows(i).Cells(21).Value <> "" Then
                    Dim buildDateString As String = DataGridView1.Rows(i).Cells(21).Value
                    buildDate = DateTime.ParseExact(buildDateString, "MM/dd/yyyy HH:mm:ss", Nothing)
                End If
                If DataGridView1.Rows(i).Cells(22).Value <> "" Then
                    Dim testingDateString As String = DataGridView1.Rows(i).Cells(22).Value
                    testingDate = DateTime.ParseExact(testingDateString, "MM/dd/yyyy HH:mm:ss", Nothing)
                End If
                If DataGridView1.Rows(i).Cells(23).Value <> "" Then
                    Dim implemDateString As String = DataGridView1.Rows(i).Cells(23).Value
                    implemDate = DateTime.ParseExact(implemDateString, "MM/dd/yyyy HH:mm:ss", Nothing)
                End If

                Dim changePhaseState As Integer = 0
                If DataGridView1.Rows(i).Cells(18).Value = "" Then
                    For j = 18 To 23 Step 1
                        DataGridView1.Rows(i).Cells(j).Style.BackColor = Color.LightGray
                        DataGridView1.Rows(i).Cells(j).Style.ForeColor = Color.LightGray
                    Next
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "Evaluation" Then
                    changePhaseState = 1
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "Endorsement to CAB" Then
                    changePhaseState = 2
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "Build/Dev" Then
                    changePhaseState = 3
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "UAT" Then
                    changePhaseState = 4
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "Prod Transport Approval" Then
                    changePhaseState = 5
                ElseIf DataGridView1.Rows(i).Cells(18).Value = "Implementation" Then
                    changePhaseState = 6
                End If

                If changePhaseState <= 2 Then 'Endorsement to CAB reached
                    'Late Review Date = Red
                    If reviewDate <> "" Then
                        If reviewDate < DateTime.Now Then
                            DataGridView1.Rows(i).Cells(19).Style.BackColor = Color.Red
                            DataGridView1.Rows(i).Cells(19).Style.ForeColor = Color.White
                        End If
                    End If
                End If

                If changePhaseState <= 3 Then 'Build/Dev reached
                    'Late Build Date = Red
                    If buildDate <> "" Then
                        If buildDate < DateTime.Now Then
                            DataGridView1.Rows(i).Cells(21).Style.BackColor = Color.Red
                            DataGridView1.Rows(i).Cells(21).Style.ForeColor = Color.White
                        End If
                    End If
                End If

                If changePhaseState <= 4 Then 'UAT reached
                    'Late Build Date = Red
                    If testingDate <> "" Then
                        If testingDate < DateTime.Now Then
                            DataGridView1.Rows(i).Cells(22).Style.BackColor = Color.Red
                            DataGridView1.Rows(i).Cells(22).Style.ForeColor = Color.White
                        End If
                    End If
                End If

                If changePhaseState <= 5 Then 'Prod Transport Approval reached
                    'Late Implem Date = Red
                    If implemDate <> "" Then
                        If implemDate < DateTime.Now Then
                            DataGridView1.Rows(i).Cells(23).Style.BackColor = Color.Red
                            DataGridView1.Rows(i).Cells(23).Style.ForeColor = Color.White
                        End If
                    End If
                End If

                If changePhaseState = 5 Then 'Implementation reached
                    'Late Implem Date = Red
                    If implemDate <> "" Then
                        If implemDate < DateTime.Now Then
                            DataGridView1.Rows(i).Cells(23).Style.BackColor = Color.Red
                            DataGridView1.Rows(i).Cells(23).Style.ForeColor = Color.White
                        End If
                    End If
                End If
            Next

        Next
    End Sub
#End Region

#Region "MOUSE DOWN RIGHT CLICK EVENT"
    Private Sub DataGridView1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.DataGridView1.HitTest(e.X, e.Y)

            If ht.Type = DataGridViewHitTestType.Cell Then

                'Behave like Left Click
                DataGridView1.CurrentCell = Me.DataGridView1(ht.ColumnIndex, ht.RowIndex)

                'Show context menu strip
                DataGridView1.ContextMenuStrip = ContextMenuStrip1
                'mnuCell.Items(0).Text = String.Format("This is the cell at {0}, {1}", ht.ColumnIndex, ht.RowIndex)
                'Trace.WriteLine(String.Format("This is the cell at {0}, {1}", ht.ColumnIndex, ht.RowIndex))
                selectedTicket = Me.DataGridView1.Rows(ht.RowIndex).Cells("TICKET").Value.ToString()

                If selectedTicket.StartsWith("SD") Then
                    selectedWorkspace = "1"
                ElseIf selectedTicket.StartsWith("PM") Then
                    selectedWorkspace = "2"
                ElseIf selectedTicket.StartsWith("RFC") Then
                    selectedWorkspace = "3"
                Else
                    selectedWorkspace = "1"
                End If

            ElseIf ht.Type = DataGridViewHitTestType.RowHeader Then
                DataGridView1.ContextMenuStrip = Nothing
                'DataGridView1.ContextMenuStrip = ContextMenuStrip1
                'mnuRow.Items(0).Text = "This is row " + ht.RowIndex.ToString()
                'Trace.WriteLine("This is row " + ht.RowIndex.ToString())

            ElseIf ht.Type = DataGridViewHitTestType.ColumnHeader Then
                DataGridView1.ContextMenuStrip = Nothing
                'DataGridView1.ContextMenuStrip = ContextMenuStrip1
                'mnuColumn.Items(0).Text = "This is col " + ht.ColumnIndex.ToString()
                'Trace.WriteLine("This is col " + ht.ColumnIndex.ToString())

            End If
        End If
    End Sub
#End Region

#Region "TOGGLE CELL MODE"

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        ToggleCellMode()
    End Sub

    Private Sub ToggleCellMode()
        If DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect Then
            DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Else
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If
    End Sub
#End Region

#Region "ADD TICKETS"
    Private Sub AddTicket()
        If selectedQueryType = "1" Then
            Dim newTicketNo As String = InputBox("Please input the ticket number", "Footprints Viewer")
            If newTicketNo = "" Then
                'Do nothing
                Exit Sub
            ElseIf newTicketNo <> " " Then
                If newTicketNo.Contains("SD ") = True Or newTicketNo.Contains("RFC ") = True Or newTicketNo.Contains("PM ") = True Then
                    File.AppendAllText(filepathTextBox.Text, vbCrLf & newTicketNo & vbCrLf)
                    CleanTickets()
                    RefreshAction(selectedQueryType)
                Else
                    MsgBox("Ticket number must follow the correct format" & vbNewLine & "i.e. SD 10000, RFC 100, PM 10", MsgBoxStyle.OkOnly, "Footprints Viewer")
                End If
            Else
                MsgBox("Ticket number must not be empty", MsgBoxStyle.OkOnly, "Footprints Viewer")
            End If
        Else
            MsgBox("You cannot add a Ticket on this Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub CleanTickets()
        Dim ticketsFile As String = filepathTextBox.Text
        If File.Exists(ticketsFile) = True Then
            Dim lines() As String
            Dim outputlines As New List(Of String)
            Dim lastLine As String = ""

            'Sort File First
            lines = File.ReadAllLines(ticketsFile)
            For Each line As String In lines
                outputlines.Add(line)
            Next
            outputlines.Sort()
            File.WriteAllLines(ticketsFile, outputlines)

            'Delete Duplicates
            outputlines.Clear()
            lines = File.ReadAllLines(ticketsFile)
            For Each line As String In lines
                If line = "" Then
                    'Do nothing
                Else
                    If lastLine = "" Then
                        lastLine = line
                        outputlines.Add(line)
                    Else
                        If line = lastLine Then
                            'Do nothing
                        Else
                            lastLine = line
                            outputlines.Add(line)
                        End If
                    End If

                End If

            Next
            outputlines.Sort()
            File.WriteAllLines(ticketsFile, outputlines)
        End If
    End Sub

    Private Sub AddTicketToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddTicketToolStripMenuItem.Click
        AddTicket()
    End Sub

    Private Sub AddTicketToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddTicketToolStripMenuItem1.Click
        AddTicket()
    End Sub
#End Region

#Region "REMOVE TICKETS"
    Private Sub RemoveTicket()
        If selectedQueryType = "1" Then
            Dim lines() As String
            Dim outputlines As New List(Of String)

            Dim result As Integer = MessageBox.Show("Are you sure you want to remove " + selectedTicket + "?", "Remove Ticket", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                ' Do nothing
            ElseIf result = DialogResult.Yes Then
                lines = File.ReadAllLines(filepathTextBox.Text)

                For Each line As String In lines
                    If line.Contains(selectedTicket) = False Then
                        outputlines.Add(line)
                    End If
                Next

                outputlines.Sort()
                File.WriteAllLines(filepathTextBox.Text, outputlines)
                RefreshAction(selectedQueryType)
            End If
        Else
            MsgBox("You cannot remove a Ticket on this Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub RemoveTicketToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveTicketToolStripMenuItem.Click
        RemoveTicket()
    End Sub

    Private Sub RemoveTicketToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveTicketToolStripMenuItem1.Click
        RemoveTicket()
    End Sub
#End Region

#Region "GO TO TICKET"
    Private Sub GoToTicketToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToTicketToolStripMenuItem.Click
        If selectedTicket <> "" Then
            'Open Browser
            Dim bmcAddress As String = "http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL=" & selectedTicket.Split(" ")(1) & "DA" & selectedWorkspace
            Process.Start(bmcAddress)
        Else
            MsgBox("Please select a Ticket from the list", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub
#End Region

#Region "GO TO JIRA"
    Private Sub GoToJIRAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToJIRAToolStripMenuItem.Click
        'Open Browser
        Dim jiraAddress As String
        If selectedTicket.Contains("SD ") = True Then
            jiraAddress = "https://jira.indra.es/browse/MITAM-0?jql=text~%22" & DateTime.Today.ToString("MMMM") & "%20" & DateTime.Today.ToString("yyyy") & "%20Functional%20Support%22"

            '"https://jira.indra.es/secure/QuickSearch.jspa?searchString=" & DateTime.Today.ToString("MMMM") & "+" & DateTime.Today.ToString("yyyy") & "+Functional+Support"



        Else
            jiraAddress = "https://jira.indra.es/secure/QuickSearch.jspa?searchString=" & selectedTicket.Replace(" ", "+")
        End If
        Process.Start(jiraAddress)
    End Sub
#End Region

#Region "UPDATE PROGRESS PERCENT"
    Private Sub DataGridView1_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        updateProgressFile(e.RowIndex)
    End Sub

    Private Sub updateProgressFile(ByVal rowIndex As Integer)
        Dim progressFile As String = "Progress.dat"
        If File.Exists(progressFile) Then
            Dim lines() As String
            Dim outputlines As New List(Of String)
            lines = File.ReadAllLines(progressFile)

            For Each line As String In lines
                Dim lineTicketNumber As String = line.Split(" ")(0) + " " + line.Split(" ")(1)
                If lineTicketNumber <> Me.DataGridView1.Rows(rowIndex).Cells("TICKET").Value.ToString() Then
                    outputlines.Add(line)
                End If
            Next

            outputlines.Sort()
            File.WriteAllLines(progressFile, outputlines)
            Try
                If Me.DataGridView1.Rows(rowIndex).Cells("PROGRESS").Value.ToString() <> "" Then
                    Using sw As StreamWriter = File.AppendText(progressFile)
                        sw.WriteLine(Me.DataGridView1.Rows(rowIndex).Cells("TICKET").Value.ToString() & " " & Me.DataGridView1.Rows(rowIndex).Cells("PROGRESS").Value.ToString())
                    End Using
                End If
            Catch nre As NullReferenceException
                'Do nothing - Text Entered is Blank                
            End Try
        Else
            'File.CreateText("Progress.dat")
            Using sw As StreamWriter = File.CreateText(progressFile)
                sw.WriteLine(Me.DataGridView1.Rows(rowIndex).Cells("TICKET").Value.ToString() & " " & Me.DataGridView1.Rows(rowIndex).Cells("PROGRESS").Value.ToString())
            End Using

        End If
    End Sub

    Private Sub markProgressAs100(ByVal status As String)
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(9).Value = status And DataGridView1.Rows(i).Cells(1).Value <> "100%" Then
                DataGridView1.Rows(i).Cells(1).Value = "100%"
                updateProgressFile(i)
            End If
        Next
    End Sub

    Private Sub MarkAllClosedAs100ProgressToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkAllClosedAs100ProgressToolStripMenuItem.Click
        markProgressAs100("Closed")
    End Sub

    Private Sub MarkAllCancelledAs100ProgressToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkAllCancelledAs100ProgressToolStripMenuItem.Click
        markProgressAs100("Cancelled")
    End Sub

    Private Sub MarkAllResolvedAs100ProgressToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkAllResolvedAs100ProgressToolStripMenuItem.Click
        markProgressAs100("Resolved")
    End Sub

    Private Sub ClearAllProgressInCurrentListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearAllProgressInCurrentListToolStripMenuItem.Click
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value <> "" Then
                DataGridView1.Rows(i).Cells(1).Value = ""
                updateProgressFile(i)
            End If
        Next
    End Sub
#End Region

#Region "EDIT IN NOTEPAD"
    Private Sub QuickEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuickEdit.Click
        'Open notepad
        If File.Exists(filepathTextBox.Text) Then
            Process.Start("notepad.exe", filepathTextBox.Text)
        Else
            Dim createFileQuestion = MessageBox.Show("This tickets file does not exist. Would you like to create it?", "Footprints", MessageBoxButtons.YesNo)
            If createFileQuestion = vbYes Then
                Using sw As StreamWriter = File.CreateText(filepathTextBox.Text)
                    'Do nothing
                End Using
                Process.Start("notepad.exe", filepathTextBox.Text)
            End If
        End If
    End Sub
#End Region

#Region "REMEMBER FILE PATH"
    Private Sub RememberFilePath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RememberFilePath.Click
        'Save file path to Application Settings
        My.Settings.FilePath = filepathTextBox.Text
        My.Settings.Save()
    End Sub
#End Region

#Region "CLEAN PROGRESS FILE"
    Private Sub CleanProgressFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CleanProgressFileToolStripMenuItem.Click
        If selectedQueryType = "1" Then
            Dim result As Integer = MessageBox.Show("Are you sure you want to clean the progress file? Cleaning the progress file means deleting the progress of every ticket which are not in your tickets list.", "Clean Progress File", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            End If

            Dim progressFile As String = "Progress.dat"
            If File.Exists(progressFile) = True Then
                Dim lines() As String
                Dim outputlines As New List(Of String)
                Dim lastLine As String = ""

                'Sort File First
                lines = File.ReadAllLines(progressFile)
                For Each line As String In lines
                    outputlines.Add(line)
                Next
                outputlines.Sort()
                File.WriteAllLines(progressFile, outputlines)

                'Delete Duplicates
                outputlines.Clear()
                lines = File.ReadAllLines(progressFile)
                For Each line As String In lines
                    If line = "" Then
                        'Do nothing
                    Else
                        'Check if Ticket is still on the list
                        Dim ticketStillOnTheList As Boolean = False
                        For i As Integer = 0 To DataGridView1.RowCount - 1
                            If line.Split(" ")(0) + " " + line.Split(" ")(1) = DataGridView1.Rows(i).Cells("TICKET").Value.ToString() Then
                                ticketStillOnTheList = True
                                Exit For
                            End If
                        Next

                        If ticketStillOnTheList = True Then
                            If lastLine = "" Then
                                lastLine = line
                                outputlines.Add(line)
                            Else
                                If line = lastLine Then
                                    'Do nothing
                                Else
                                    lastLine = line
                                    outputlines.Add(line)
                                End If
                            End If
                        Else
                            'Do nothing
                        End If

                    End If
                Next
                outputlines.Sort()
                File.WriteAllLines(progressFile, outputlines)
            End If
        Else
            MsgBox("You cannot clean the progress file while using this Query Type.", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub
#End Region

#Region "FIND"
    Private Sub FindStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindStripMenuItem.Click
        Dim searchIndex = 0
        Dim findString As String = InputBox("Find what?", "Footprints Viewer")

        If findString <> "" Then
            'Reset Search String for Find NExy
            currentFindString = findString
            currentSearchIndex = 0

            DataGridView1.ClearSelection()
            selectedTicket = ""
            For Each row As DataGridViewRow In DataGridView1.Rows
                For Each cell As DataGridViewCell In row.Cells
                    If cell.Value.ToString.Contains(findString) Then
                        If searchIndex = currentSearchIndex Then
                            'If currentSearchIndex has a value of 0, it would select the first match, second match for a value of 1, etc.
                            'This is the cell we want to select
                            cell.Selected = True
                            selectedTicket = Me.DataGridView1.Rows(cell.RowIndex).Cells("TICKET").Value.ToString()
                            'Focus cell
                            DataGridView1.CurrentCell = DataGridView1.Rows(row.Index).Cells(cell.ColumnIndex)
                        End If
                        'Yellow background for all matches
                        'cell.Style.BackColor = Color.Yellow
                        searchIndex += 1
                    End If
                Next
            Next

            If selectedTicket = "" Then
                MsgBox("""" & findString & """ not found", MsgBoxStyle.OkOnly, "Footprints Viewer")
            End If

        End If
    End Sub

    Private Sub FindNextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindNextToolStripMenuItem.Click
        Dim searchIndex As Integer = 0

        If currentFindString <> "" Then
            'Increment for next result
            currentSearchIndex += 1

            DataGridView1.ClearSelection()
            selectedTicket = ""
            For Each row As DataGridViewRow In DataGridView1.Rows
                For Each cell As DataGridViewCell In row.Cells
                    If cell.Value.ToString.Contains(currentFindString) Then
                        If searchIndex = currentSearchIndex Then
                            'If currentSearchIndex has a value of 0, it would select the first match, second match for a value of 1, etc.
                            'This is the cell we want to select
                            cell.Selected = True
                            selectedTicket = Me.DataGridView1.Rows(cell.RowIndex).Cells("TICKET").Value.ToString()
                            'Focus cell
                            DataGridView1.CurrentCell = DataGridView1.Rows(row.Index).Cells(cell.ColumnIndex)
                        End If
                        'Yellow background for all matches
                        'cell.Style.BackColor = Color.Yellow
                        searchIndex += 1
                    End If
                Next
            Next

            ' If there are no more results, redo the search from the beginning
            If selectedTicket = "" Then
                searchIndex = 0
                currentSearchIndex = 0
                DataGridView1.ClearSelection()
                selectedTicket = ""
                For Each row As DataGridViewRow In DataGridView1.Rows
                    For Each cell As DataGridViewCell In row.Cells
                        If cell.Value.ToString.Contains(currentFindString) Then
                            If searchIndex = currentSearchIndex Then
                                'If currentSearchIndex has a value of 0, it would select the first match, second match for a value of 1, etc.
                                'This is the cell we want to select
                                cell.Selected = True
                                selectedTicket = Me.DataGridView1.Rows(cell.RowIndex).Cells("TICKET").Value.ToString()
                                'Focus cell
                                DataGridView1.CurrentCell = DataGridView1.Rows(row.Index).Cells(cell.ColumnIndex)
                            End If
                            'Yellow background for all matches
                            'cell.Style.BackColor = Color.Yellow
                            searchIndex += 1
                        End If
                    Next
                Next
            End If
        End If

    End Sub
#End Region

#Region "OPEN FILE DIALOG"
    Private Sub OpenFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenFileToolStripMenuItem.Click
        Dim openFileDialog As New OpenFileDialog

        openFileDialog.Filter = "Text Files|*.txt"
        openFileDialog.Title = "Select a Tickets File"
        If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filepathTextBox.Text = openFileDialog.FileName()
        End If

    End Sub
#End Region

#Region "VIEW TICKET IN BMC FOOTPRINTS"
    Private Sub ViewATicketInBMCFootprintsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewATicketInBMCFootprintsToolStripMenuItem.Click
        Dim inputTicketNumber As String = InputBox("What is the Ticket Number?", "Footprints Viewer")
        Dim inputWorkspace As String = ""

        Try
            If inputTicketNumber = "" Then
                'Do nothing
                Exit Sub
            ElseIf inputTicketNumber <> " " Then
                If inputTicketNumber.StartsWith("SD") Then
                    inputWorkspace = "1"
                ElseIf inputTicketNumber.StartsWith("PM") Then
                    inputWorkspace = "2"
                ElseIf inputTicketNumber.StartsWith("RFC") Then
                    inputWorkspace = "3"
                Else
                    inputWorkspace = ""
                End If

                If inputWorkspace <> "" Then
                    'Open Browser
                    Dim bmcAddress As String = "http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL=" & inputTicketNumber.Split(" ")(1) & "DA" & inputWorkspace
                    Process.Start(bmcAddress)
                Else
                    MsgBox("Ticket number must follow the correct format" & vbNewLine & "i.e. SD 10000, RFC 100, PM 10", MsgBoxStyle.OkOnly, "Footprints Viewer")
                End If
            Else
                MsgBox("Ticket number must not be empty", MsgBoxStyle.OkOnly, "Footprints Viewer")
            End If
        Catch ex As Exception
            MsgBox("Ticket number must follow the correct format" & vbNewLine & "i.e. SD 10000, RFC 100, PM 10", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End Try
    End Sub
#End Region

#Region "MIGRATE PROGRESS FILE"
    Private Sub MigrateAProgressFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MigrateAProgressFileToolStripMenuItem.Click
        Dim openFileDialog As New OpenFileDialog
        Dim oldProgressFile As String = ""
        Dim newProgressFile As String = "Progress.dat"
        Dim lines() As String
        Dim outputlines As New List(Of String)
        Dim lastLine As String = ""

        openFileDialog.Filter = "Text Files|*.txt|DAT Files|*.dat"
        openFileDialog.Title = "Select a Tickets File"
        If openFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            oldProgressFile = openFileDialog.FileName()

            'Sort File First
            lines = File.ReadAllLines(oldProgressFile)
            For Each line As String In lines
                outputlines.Add(line)
            Next
            outputlines.Sort()
            File.WriteAllLines(newProgressFile, outputlines)

            MsgBox("Successfully migrated progress file!" & vbNewLine & "Please click refresh to see the changes in your progress.", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If

    End Sub
#End Region

#Region "QUERY TYPE COMBO BOX EVENT"
    Private Sub QueryTypeToolStripComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QueryTypeToolStripComboBox.SelectedIndexChanged
        If QueryTypeToolStripComboBox.SelectedIndex = 0 Then
            selectedQueryType = "1"
            filepathToolStripLabel.Visible = True
            filepathTextBox.Visible = True
            RememberFilePath.Visible = True
            QuickEdit.Visible = True
            SubcategoryToolStripComboBox.Visible = False
            SubcategoryToolStripLabel.Visible = False
            UsernameLabel.Visible = False
            UsernameToolStripTextBox.Visible = False
            DescriptionToolStripLabel.Visible = False
            DescriptionToolStripTextBox.Visible = False

            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked

        ElseIf QueryTypeToolStripComboBox.SelectedIndex = 1 Then
            selectedQueryType = "2"
            filepathToolStripLabel.Visible = False
            filepathTextBox.Visible = False
            RememberFilePath.Visible = False
            QuickEdit.Visible = False
            SubcategoryToolStripComboBox.Visible = True
            SubcategoryToolStripLabel.Visible = True
            UsernameLabel.Visible = True
            UsernameToolStripTextBox.Visible = True
            DescriptionToolStripLabel.Visible = True
            DescriptionToolStripTextBox.Visible = True
            createSubCategoryItems()
        End If
    End Sub
#End Region

#Region "TOGGLE COLUMNS"
    Private Sub ToggleColumns()
        If Me.DataGridView1.ColumnCount > 0 Then
            If TicketSummaryToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("TICKET_SUMMARY").Visible = True
            Else
                Me.DataGridView1.Columns("TICKET_SUMMARY").Visible = False
            End If

            If ProgressToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("PROGRESS").Visible = True
            Else
                Me.DataGridView1.Columns("PROGRESS").Visible = False
            End If

            If CreatedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("CREATED").Visible = True
            Else
                Me.DataGridView1.Columns("CREATED").Visible = False
            End If

            If DueDateToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("DUE_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("DUE_DATE").Visible = False
            End If

            If CategoryToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("CATEGORY").Visible = True
            Else
                Me.DataGridView1.Columns("CATEGORY").Visible = False
            End If

            If SubCategoryToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("SUB_CATEGORY").Visible = True
            Else
                Me.DataGridView1.Columns("SUB_CATEGORY").Visible = False
            End If

            If RequestorToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("REQUESTOR").Visible = True
            Else
                Me.DataGridView1.Columns("REQUESTOR").Visible = False
            End If

            If TicketToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("TICKET").Visible = True
            Else
                Me.DataGridView1.Columns("TICKET").Visible = False
            End If

            If DescriptionToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("DESCRIPTION").Visible = True
            Else
                Me.DataGridView1.Columns("DESCRIPTION").Visible = False
            End If

            If StatusToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("STATUS").Visible = True
            Else
                Me.DataGridView1.Columns("STATUS").Visible = False
            End If

            If RemarksToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("REMARKS").Visible = True
            Else
                Me.DataGridView1.Columns("REMARKS").Visible = False
            End If

            If ClassificationToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("CLASSIFICATION").Visible = True
            Else
                Me.DataGridView1.Columns("CLASSIFICATION").Visible = False
            End If

            If AssigneesToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("ASSIGNEES").Visible = True
            Else
                Me.DataGridView1.Columns("ASSIGNEES").Visible = False
            End If

            If ModifiedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("MODIFIED").Visible = True
            Else
                Me.DataGridView1.Columns("MODIFIED").Visible = False
            End If

            If ChangePhaseToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("CHANGE_PHASE").Visible = True
            Else
                Me.DataGridView1.Columns("CHANGE_PHASE").Visible = False
            End If

            If DateReviewedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("REVIEW_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("REVIEW_DATE").Visible = False
            End If

            If DateApprovedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("AUTHORIZATION_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("AUTHORIZATION_DATE").Visible = False
            End If

            If BuildDevDateToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("BUILD_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("BUILD_DATE").Visible = False
            End If

            If TestingDateToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("TESTING_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("TESTING_DATE").Visible = False
            End If

            If DateImplementedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("IMPLEM_DATE").Visible = True
            Else
                Me.DataGridView1.Columns("IMPLEM_DATE").Visible = False
            End If

            If ResolvedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("RESOLVED").Visible = True
            Else
                Me.DataGridView1.Columns("RESOLVED").Visible = False
            End If

            If KBToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("KB").Visible = True
            Else
                Me.DataGridView1.Columns("KB").Visible = False
            End If

            If DateLinkedToKBToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("DATE_LINKED_TO_KB").Visible = True
            Else
                Me.DataGridView1.Columns("DATE_LINKED_TO_KB").Visible = False
            End If

            If TCodeToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("T_CODE").Visible = True
            Else
                Me.DataGridView1.Columns("T_CODE").Visible = False
            End If

            If LastUpdatedToolStripMenuItem.CheckState = CheckState.Checked Then
                Me.DataGridView1.Columns("LAST_UPDATED").Visible = True
            Else
                Me.DataGridView1.Columns("LAST_UPDATED").Visible = False
            End If

        End If
    End Sub
#End Region

#Region "ACTION EVENTS - COLUMNS"
    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TicketSummaryToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub ProgressToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub CreatedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreatedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DueDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DueDateToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub SubCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubCategoryToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub RequestorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RequestorToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub TicketToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TicketToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DescriptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DescriptionToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub StatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub RemarksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemarksToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub ClassificationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClassificationToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub AssigneesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssigneesToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub ModifiedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModifiedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub ResolvedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResolvedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub KBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KBToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DateLinkedToKBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateLinkedToKBToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub TCodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCodeToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub BuildDevDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BuildDevDateToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub TestingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestingDateToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DateImplementedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateImplementedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DateReviewedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateReviewedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub DateApprovedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateApprovedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub ChangePhaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangePhaseToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub LastUpdatedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastUpdatedToolStripMenuItem.Click
        ToggleColumns()
    End Sub

    Private Sub AllColumnsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllColumnsToolStripMenuItem.Click
        TicketSummaryToolStripMenuItem.Checked = True
        ProgressToolStripMenuItem.Checked = True
        CreatedToolStripMenuItem.Checked = True
        DueDateToolStripMenuItem.Checked = True
        CategoryToolStripMenuItem.Checked = True
        SubCategoryToolStripMenuItem.Checked = True
        RequestorToolStripMenuItem.Checked = True
        TicketToolStripMenuItem.Checked = True
        DescriptionToolStripMenuItem.Checked = True
        StatusToolStripMenuItem.Checked = True
        RemarksToolStripMenuItem.Checked = True
        ClassificationToolStripMenuItem.Checked = True
        AssigneesToolStripMenuItem.Checked = True
        ModifiedToolStripMenuItem.Checked = True
        ResolvedToolStripMenuItem.Checked = True
        KBToolStripMenuItem.Checked = True
        DateLinkedToKBToolStripMenuItem.Checked = True
        TCodeToolStripMenuItem.Checked = True
        ChangePhaseToolStripMenuItem.Checked = True
        DateReviewedToolStripMenuItem.Checked = True
        DateApprovedToolStripMenuItem.Checked = True
        BuildDevDateToolStripMenuItem.Checked = True
        TestingDateToolStripMenuItem.Checked = True
        DateImplementedToolStripMenuItem.Checked = True
        LastUpdatedToolStripMenuItem.Checked = True
        ToggleColumns()
    End Sub

    Private Sub WeeklyStatusReportToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeeklyStatusReportToolStripMenuItem.Click
        TicketSummaryToolStripMenuItem.Checked = True
        ProgressToolStripMenuItem.Checked = True
        CreatedToolStripMenuItem.Checked = True
        DueDateToolStripMenuItem.Checked = True
        CategoryToolStripMenuItem.Checked = True
        SubCategoryToolStripMenuItem.Checked = True
        RequestorToolStripMenuItem.Checked = True
        TicketToolStripMenuItem.Checked = True
        DescriptionToolStripMenuItem.Checked = True
        StatusToolStripMenuItem.Checked = True
        RemarksToolStripMenuItem.Checked = True
        ClassificationToolStripMenuItem.Checked = False
        AssigneesToolStripMenuItem.Checked = False
        ModifiedToolStripMenuItem.Checked = False
        ResolvedToolStripMenuItem.Checked = False
        KBToolStripMenuItem.Checked = False
        DateLinkedToKBToolStripMenuItem.Checked = False
        TCodeToolStripMenuItem.Checked = False
        ChangePhaseToolStripMenuItem.Checked = False
        DateReviewedToolStripMenuItem.Checked = False
        DateApprovedToolStripMenuItem.Checked = False
        BuildDevDateToolStripMenuItem.Checked = False
        TestingDateToolStripMenuItem.Checked = False
        DateImplementedToolStripMenuItem.Checked = False
        LastUpdatedToolStripMenuItem.Checked = False
        ToggleColumns()
    End Sub

    Private Sub OverdueDueTodayDueTomorrowToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OverdueDueTodayDueTomorrowToolStripMenuItem.Click
        TicketSummaryToolStripMenuItem.Checked = False
        ProgressToolStripMenuItem.Checked = False
        CreatedToolStripMenuItem.Checked = False
        DueDateToolStripMenuItem.Checked = True
        CategoryToolStripMenuItem.Checked = True
        SubCategoryToolStripMenuItem.Checked = True
        RequestorToolStripMenuItem.Checked = True
        TicketToolStripMenuItem.Checked = True
        DescriptionToolStripMenuItem.Checked = True
        StatusToolStripMenuItem.Checked = True
        RemarksToolStripMenuItem.Checked = True
        ClassificationToolStripMenuItem.Checked = True
        AssigneesToolStripMenuItem.Checked = True
        ModifiedToolStripMenuItem.Checked = False
        ResolvedToolStripMenuItem.Checked = False
        KBToolStripMenuItem.Checked = False
        DateLinkedToKBToolStripMenuItem.Checked = False
        TCodeToolStripMenuItem.Checked = False
        ChangePhaseToolStripMenuItem.Checked = False
        DateReviewedToolStripMenuItem.Checked = False
        DateApprovedToolStripMenuItem.Checked = False
        BuildDevDateToolStripMenuItem.Checked = False
        TestingDateToolStripMenuItem.Checked = False
        DateImplementedToolStripMenuItem.Checked = False
        LastUpdatedToolStripMenuItem.Checked = False
        ToggleColumns()
    End Sub

    Private Sub DaysNoUpdatesToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DaysNoUpdatesToolStripMenuItem.Click
        TicketSummaryToolStripMenuItem.Checked = False
        ProgressToolStripMenuItem.Checked = False
        CreatedToolStripMenuItem.Checked = False
        DueDateToolStripMenuItem.Checked = False
        CategoryToolStripMenuItem.Checked = False
        SubCategoryToolStripMenuItem.Checked = False
        RequestorToolStripMenuItem.Checked = False
        TicketToolStripMenuItem.Checked = True
        DescriptionToolStripMenuItem.Checked = False
        StatusToolStripMenuItem.Checked = False
        RemarksToolStripMenuItem.Checked = True
        ClassificationToolStripMenuItem.Checked = True
        AssigneesToolStripMenuItem.Checked = True
        ModifiedToolStripMenuItem.Checked = True
        ResolvedToolStripMenuItem.Checked = False
        KBToolStripMenuItem.Checked = False
        DateLinkedToKBToolStripMenuItem.Checked = False
        TCodeToolStripMenuItem.Checked = False
        ChangePhaseToolStripMenuItem.Checked = False
        DateReviewedToolStripMenuItem.Checked = False
        DateApprovedToolStripMenuItem.Checked = False
        BuildDevDateToolStripMenuItem.Checked = False
        TestingDateToolStripMenuItem.Checked = False
        DateImplementedToolStripMenuItem.Checked = False
        LastUpdatedToolStripMenuItem.Checked = True
        ToggleColumns()
    End Sub

    Private Sub ChangePhaseAndTimelineToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangePhaseAndTimelineToolStripMenuItem.Click
        TicketSummaryToolStripMenuItem.Checked = False
        ProgressToolStripMenuItem.Checked = False
        CreatedToolStripMenuItem.Checked = False
        DueDateToolStripMenuItem.Checked = False
        CategoryToolStripMenuItem.Checked = False
        SubCategoryToolStripMenuItem.Checked = False
        RequestorToolStripMenuItem.Checked = False
        TicketToolStripMenuItem.Checked = True
        DescriptionToolStripMenuItem.Checked = False
        StatusToolStripMenuItem.Checked = False
        RemarksToolStripMenuItem.Checked = False
        ClassificationToolStripMenuItem.Checked = False
        AssigneesToolStripMenuItem.Checked = False
        ModifiedToolStripMenuItem.Checked = False
        ResolvedToolStripMenuItem.Checked = False
        KBToolStripMenuItem.Checked = False
        DateLinkedToKBToolStripMenuItem.Checked = False
        TCodeToolStripMenuItem.Checked = False
        ChangePhaseToolStripMenuItem.Checked = True
        DateReviewedToolStripMenuItem.Checked = True
        DateApprovedToolStripMenuItem.Checked = True
        BuildDevDateToolStripMenuItem.Checked = True
        TestingDateToolStripMenuItem.Checked = True
        DateImplementedToolStripMenuItem.Checked = True
        LastUpdatedToolStripMenuItem.Checked = False
        ToggleColumns()
    End Sub
#End Region

#Region "SAVE COLUMN LAYOUT"
    Private Sub SaveColumnLayoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveColumnLayoutToolStripMenuItem.Click

        If DataGridView1.ColumnCount > 0 Then
            ' Save columns order into application settings
            If DataGridView1.ColumnCount > 0 Then
                ' Save Column Order
                If My.Settings.ColumnOrder Is Nothing Then
                    My.Settings.ColumnOrder = New System.Collections.Specialized.StringCollection
                End If
                My.Settings.ColumnOrder.Clear()
                Dim dc As DataGridViewColumn
                For Each dc In DataGridView1.Columns
                    My.Settings.ColumnOrder.Add(dc.DisplayIndex.ToString())
                Next
                ' End - Save Column Order

                ' Save Column Visibility
                My.Settings.TicketSummaryVisibility = TicketSummaryToolStripMenuItem.CheckState()
                My.Settings.ProgressVisibility = ProgressToolStripMenuItem.CheckState()
                My.Settings.CreatedVisibility = CreatedToolStripMenuItem.CheckState()
                My.Settings.DueDateVisibility = DueDateToolStripMenuItem.CheckState()
                My.Settings.CategoryVisibility = CategoryToolStripMenuItem.CheckState()
                My.Settings.SubCategoryVisibility = SubCategoryToolStripMenuItem.CheckState()
                My.Settings.RequestorVisibility = RequestorToolStripMenuItem.CheckState()
                My.Settings.TicketVisibility = TicketToolStripMenuItem.CheckState()
                My.Settings.DescriptionVisibility = DescriptionToolStripMenuItem.CheckState()
                My.Settings.StatusVisibility = StatusToolStripMenuItem.CheckState()
                My.Settings.RemarksVisibility = RemarksToolStripMenuItem.CheckState()
                My.Settings.ClassificationVisibility = ClassificationToolStripMenuItem.CheckState()
                My.Settings.AssigneesVisibility = AssigneesToolStripMenuItem.CheckState()
                My.Settings.ModifiedVisibility = ModifiedToolStripMenuItem.CheckState()
                My.Settings.ChangePhaseVisibility = ChangePhaseToolStripMenuItem.CheckState
                My.Settings.ReviewDateVisibility = DateReviewedToolStripMenuItem.CheckState
                My.Settings.AuthDateVisibility = DateApprovedToolStripMenuItem.CheckState
                My.Settings.BuildDateVisibility = BuildDevDateToolStripMenuItem.CheckState
                My.Settings.TestingDateVisibility = TestingDateToolStripMenuItem.CheckState
                My.Settings.ImplemDateVisibility = DateImplementedToolStripMenuItem.CheckState
                My.Settings.ResolvedVisibility = ResolvedToolStripMenuItem.CheckState()
                My.Settings.KBVisibility = KBToolStripMenuItem.CheckState()
                My.Settings.DateLinkedToKBVisibility = DateLinkedToKBToolStripMenuItem.CheckState()
                My.Settings.TCodeVisibility = TCodeToolStripMenuItem.CheckState()
                My.Settings.LastUpdatedVisibility = LastUpdatedToolStripMenuItem.CheckState()
                ' End - Save Column Visibility

                My.Settings.Save()
            Else
                MsgBox("There are no columns to save", MsgBoxStyle.OkOnly, "Footprints Viewer")
            End If
        End If

    End Sub

    Private Sub RestoreColumnLayoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestoreColumnLayoutToolStripMenuItem.Click
        Try
            If DataGridView1.ColumnCount > 0 Then
                If My.Settings.ColumnOrder.Count > -1 Then
                    ' Load columns order from application settings
                    Dim i As Integer

                    For i = 0 To My.Settings.ColumnOrder.Count - 1
                        DataGridView1.Columns(i).DisplayIndex = CInt(My.Settings.ColumnOrder(i))
                    Next

                    ' Load column visibility
                    TicketSummaryToolStripMenuItem.CheckState = My.Settings.TicketSummaryVisibility
                    ProgressToolStripMenuItem.CheckState = My.Settings.ProgressVisibility
                    CreatedToolStripMenuItem.CheckState = My.Settings.CreatedVisibility
                    DueDateToolStripMenuItem.CheckState = My.Settings.DueDateVisibility
                    CategoryToolStripMenuItem.CheckState = My.Settings.CategoryVisibility
                    SubCategoryToolStripMenuItem.CheckState = My.Settings.SubCategoryVisibility
                    RequestorToolStripMenuItem.CheckState = My.Settings.RequestorVisibility
                    TicketToolStripMenuItem.CheckState = My.Settings.TicketVisibility
                    DescriptionToolStripMenuItem.CheckState = My.Settings.DescriptionVisibility
                    StatusToolStripMenuItem.CheckState = My.Settings.StatusVisibility
                    RemarksToolStripMenuItem.CheckState = My.Settings.RemarksVisibility
                    ClassificationToolStripMenuItem.CheckState = My.Settings.ClassificationVisibility
                    AssigneesToolStripMenuItem.CheckState = My.Settings.AssigneesVisibility
                    ModifiedToolStripMenuItem.CheckState = My.Settings.ModifiedVisibility
                    ChangePhaseToolStripMenuItem.CheckState = My.Settings.ChangePhaseVisibility
                    DateReviewedToolStripMenuItem.CheckState = My.Settings.ReviewDateVisibility
                    DateApprovedToolStripMenuItem.CheckState = My.Settings.AuthDateVisibility
                    BuildDevDateToolStripMenuItem.CheckState = My.Settings.BuildDateVisibility
                    TestingDateToolStripMenuItem.CheckState = My.Settings.TestingDateVisibility
                    DateImplementedToolStripMenuItem.CheckState = My.Settings.ImplemDateVisibility
                    ResolvedToolStripMenuItem.CheckState = My.Settings.ResolvedVisibility
                    KBToolStripMenuItem.CheckState = My.Settings.KBVisibility
                    DateLinkedToKBToolStripMenuItem.CheckState = My.Settings.DateLinkedToKBVisibility
                    TCodeToolStripMenuItem.CheckState = My.Settings.TCodeVisibility
                    LastUpdatedToolStripMenuItem.CheckState = My.Settings.LastUpdatedVisibility

                    'Refresh columns
                    ToggleColumns()
                End If
            End If
        Catch nre As NullReferenceException
            MsgBox("No column layouts were previously saved", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End Try
    End Sub
#End Region

#Region "SET DATE FILTER"

    Private Sub SetADateFilterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetADateFilterToolStripMenuItem.Click
        Dim dateModifiedFilter As New DateModifiedFilterForm()

        If (dateModifiedFilter.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            dateModifiedFromFilter = dateModifiedFilter.DateModifiedFromTimePicker.Value.ToString("yyyy-MM-dd 00:00:00")
            dateModifiedToFilter = dateModifiedFilter.DateModifiedToTimePicker.Value.ToString("yyyy-MM-dd 23:59:59")
            SetADateFilterToolStripMenuItem.CheckState = CheckState.Checked

            RefreshAction(selectedQueryType)
        End If
    End Sub

    Private Sub ClearDateFilterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearDateFilterToolStripMenuItem.Click
        dateModifiedFromFilter = ""
        dateModifiedToFilter = ""
        SetADateFilterToolStripMenuItem.CheckState = CheckState.Unchecked
        RefreshAction(selectedQueryType)
    End Sub
#End Region

#Region "SET STATUS FILTER"
    Private Sub ToggleTicketStatusFilter()

        If AssignedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusAssigned = True
        Else
            ticketStatusAssigned = False
        End If

        If AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusAccepted = True
        Else
            ticketStatusAccepted = False
        End If

        If PendingStatusToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusPending = True
        Else
            ticketStatusPending = False
        End If

        If ResolvedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusResolved = True
        Else
            ticketStatusResolved = False
        End If

        If ClosedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusClosed = True
        Else
            ticketStatusClosed = False
        End If

        If CancelledToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketStatusCancelled = True
        Else
            ticketStatusCancelled = False
        End If

        MsgBox("Refresh the List to see the Changes.", MsgBoxStyle.OkOnly, "Footprints Viewer")

    End Sub

    Private Sub AssignedStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssignedStatusToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            ToggleTicketStatusFilter()
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub AcceptedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcceptedStatusToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            ToggleTicketStatusFilter()
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub PendingStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PendingStatusToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            ToggleTicketStatusFilter()
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub ResolvedStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResolvedStatusToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            If ResolvedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
                Dim result As Integer = MessageBox.Show("Are you sure you want to turn on the Resolved status Filter? There could be a lot.", "Warning", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    ResolvedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
                    ticketStatusResolved = False
                Else
                    ToggleTicketStatusFilter()
                End If
            Else
                ToggleTicketStatusFilter()
            End If
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub ClosedStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClosedStatusToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            If ClosedStatusToolStripMenuItem.CheckState = CheckState.Checked Then
                Dim result As Integer = MessageBox.Show("Are you sure you want to turn on the Closed status Filter? There could be a lot.", "Warning", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
                    ticketStatusClosed = False
                Else
                    ToggleTicketStatusFilter()
                End If
            Else
                ToggleTicketStatusFilter()
            End If
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub

    Private Sub CancelledToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelledToolStripMenuItem.Click
        If selectedQueryType = 2 Then
            If CancelledToolStripMenuItem.CheckState = CheckState.Checked Then
                Dim result As Integer = MessageBox.Show("Are you sure you want to turn on the Cancelled status Filter? There could be a lot.", "Warning", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
                    ticketStatusCancelled = False
                Else
                    ToggleTicketStatusFilter()
                End If
            Else
                ToggleTicketStatusFilter()
            End If
        Else
            ticketStatusAssigned = True
            ticketStatusAccepted = True
            ticketStatusPending = True
            ticketStatusResolved = False
            ticketStatusClosed = False
            ticketStatusCancelled = False
            AcceptedStatusToolStripMenuItem.CheckState = CheckState.Checked
            PendingStatusToolStripMenuItem.CheckState = CheckState.Checked
            ResolvedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            ClosedStatusToolStripMenuItem.CheckState = CheckState.Unchecked
            CancelledToolStripMenuItem.CheckState = CheckState.Unchecked
            MsgBox("Changing the Ticket Status is only available for the Filtered Query Type", MsgBoxStyle.OkOnly, "Footprints Viewer")
        End If
    End Sub
#End Region

#Region "TICKET TYPE FILTER"
    Private Sub RFCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RFCToolStripMenuItem.Click
        toggleTicketTypeFilter()
    End Sub

    Private Sub SDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SDToolStripMenuItem.Click
        toggleTicketTypeFilter()
    End Sub

    Private Sub PMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PMToolStripMenuItem.Click
        toggleTicketTypeFilter()
    End Sub

    Private Sub toggleTicketTypeFilter()
        If RFCToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketTypeRFC = True
        Else
            ticketTypeRFC = False
        End If

        If SDToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketTypeSD = True
        Else
            ticketTypeSD = False
        End If

        If PMToolStripMenuItem.CheckState = CheckState.Checked Then
            ticketTypePM = True
        Else
            ticketTypePM = False
        End If
        MsgBox("Refresh the List to see the Changes.", MsgBoxStyle.OkOnly, "Footprints Viewer")
    End Sub
#End Region

#Region "CHANGE PHASE FILTER"
    Private Sub EvaluationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EvaluationToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub
    Private Sub EndorsementToCABToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndorsementToCABToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub

    Private Sub BuildDevToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BuildDevToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub

    Private Sub UATToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UATToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub

    Private Sub ProdTransportApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdTransportApprovalToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub

    Private Sub ImplementationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImplementationToolStripMenuItem.Click
        toggleChangePhaseFilter()
    End Sub

    Private Sub toggleChangePhaseFilter()
        If EvaluationToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterEvaluation = True
        Else
            changePhaseFilterEvaluation = False
        End If

        If EndorsementToCABToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterEndorsementToCab = True
        Else
            changePhaseFilterEndorsementToCab = False
        End If

        If BuildDevToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterBuildDev = True
        Else
            changePhaseFilterBuildDev = False
        End If

        If UATToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterUAT = True
        Else
            changePhaseFilterUAT = False
        End If

        If ProdTransportApprovalToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterProdTransportApproval = True
        Else
            changePhaseFilterProdTransportApproval = False
        End If

        If ImplementationToolStripMenuItem.CheckState = CheckState.Checked Then
            changePhaseFilterImplementation = True
        Else
            changePhaseFilterImplementation = False
        End If
        MsgBox("Refresh the List to see the Changes.", MsgBoxStyle.OkOnly, "Footprints Viewer")
    End Sub
#End Region

End Class
