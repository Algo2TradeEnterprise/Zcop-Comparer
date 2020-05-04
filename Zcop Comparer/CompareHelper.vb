Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class CompareHelper
    Implements IDisposable

#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event HeartbeatSub(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnHeartbeatSub(ByVal msg As String)
        RaiseEvent HeartbeatSub(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _oldFilePath As String
    Private ReadOnly _newFilePath As String
    Private ReadOnly _mappingFilePath As String
    Private ReadOnly _fileSchema As Dictionary(Of String, String)
    Private ReadOnly _mappingFileSchema As Dictionary(Of String, String)

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal oldFilePath As String, ByVal newFilePath As String, ByVal mappingFilePath As String)
        _cts = canceller
        _oldFilePath = oldFilePath
        _newFilePath = newFilePath
        _mappingFilePath = mappingFilePath

        _fileSchema = New Dictionary(Of String, String) From
                          {{"Emp No", "EMP_CODE"},
                           {"Field1", "PROJECT_ACQUIRED_SKILL"},
                           {"Field2", "CERTIFIED_SKILL"},
                           {"Field3", "TRAINING_SKILLS"}}

        _mappingFileSchema = New Dictionary(Of String, String) From
                            {{"WFT Practice", "WFT Practice"},
                            {"WFT Skill Level", "WFT Skill Level"},
                            {"WFT Skill Bucket", "WFT Skill Bucket"},
                            {"WFT Weightage", "WFT Weightage"},
                            {"WFT Subskills", "WFT Subskills"}}
    End Sub

    Public Async Function StartCompareAsync() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        If Not File.Exists(_oldFilePath) Then Throw New ApplicationException("Old File is not available at given path")
        If Not File.Exists(_newFilePath) Then Throw New ApplicationException("New File is not available at given path")
        If Not File.Exists(_mappingFilePath) Then Throw New ApplicationException("Mapping File is not available at given path")

        Dim empScoreDetails As Dictionary(Of String, ScoreDetails) = Nothing
        ReadRequiedDataFromFile(_oldFilePath, "Old", empScoreDetails)
        ReadRequiedDataFromFile(_newFilePath, "New", empScoreDetails)
        Dim mappingSkills As List(Of String) = ReadRequiedMappingFile(_mappingFilePath)
        'Start Comparison
        If empScoreDetails IsNot Nothing AndAlso empScoreDetails.Count > 0 Then
            OnHeartbeat("Comparing score")
            Dim counter As Integer = 0
            For Each runningEmp In empScoreDetails
                _cts.Token.ThrowIfCancellationRequested()
                counter += 1
                OnHeartbeatSub(String.Format("Comparing scores # {0}/{1}", counter, empScoreDetails.Count))
                CompareAllScores(runningEmp.Value, mappingSkills)
            Next
            OnHeartbeatSub("")

            _cts.Token.ThrowIfCancellationRequested()
            OnHeartbeat("Writing data to memory")
            Dim columnCount As Integer = (_fileSchema.Count - 1) * 3
            Dim output(empScoreDetails.Count, columnCount + 3)
            'Writing Columns
            Dim clmCtr As Integer = 0
            output(0, clmCtr) = _fileSchema("Emp No")
            For fieldCount As Integer = 1 To _fileSchema.Count - 1
                _cts.Token.ThrowIfCancellationRequested()
                Dim fieldName As String = _fileSchema.ElementAt(fieldCount).Value
                clmCtr += 1
                output(0, clmCtr) = String.Format("OLD {0}", fieldName.ToUpper)
                clmCtr += 1
                output(0, clmCtr) = String.Format("NEW {0}", fieldName.ToUpper)
                clmCtr += 1
                output(0, clmCtr) = String.Format("DELTA {0}", fieldName.ToUpper)
            Next
            clmCtr += 1
            output(0, clmCtr) = "Manually Updated"
            clmCtr += 1
            output(0, clmCtr) = "System Updated"
            clmCtr += 1
            output(0, clmCtr) = "Overall Updated"

            'Writing Data
            Dim rowCtr As Integer = 0
            For Each runningEmp In empScoreDetails
                _cts.Token.ThrowIfCancellationRequested()
                rowCtr += 1
                OnHeartbeatSub(String.Format("Writing data to memory # {0}/{1}", rowCtr - 1, empScoreDetails.Count))
                Dim columnCtr As Integer = 0

                output(rowCtr, columnCtr) = runningEmp.Value.EmpNo
                Dim manualUpdate As Boolean = False
                Dim systemUpdate As Boolean = False
                If runningEmp.Value.Fields IsNot Nothing AndAlso runningEmp.Value.Fields.Count > 0 Then
                    For Each runningField In runningEmp.Value.Fields
                        If runningField.Value.Changes IsNot Nothing Then
                            columnCtr += 1
                            output(rowCtr, columnCtr) = GetRawSkillString(runningField.Value.OldFileData, mappingSkills)
                            columnCtr += 1
                            output(rowCtr, columnCtr) = GetRawSkillString(runningField.Value.NewFileData, mappingSkills)
                            columnCtr += 1
                            output(rowCtr, columnCtr) = runningField.Value.Changes.OverallChange
                            If runningField.Value.Changes.OverallChange IsNot Nothing Then
                                If runningField.Key.ToUpper = "FIELD1" Then
                                    manualUpdate = runningField.Value.Changes.OverallChange.Contains("(+)")
                                Else
                                    systemUpdate = runningField.Value.Changes.OverallChange.Contains("(+)")
                                End If
                            End If
                        Else
                                columnCtr += 1
                            output(rowCtr, columnCtr) = GetRawSkillString(runningField.Value.OldFileData, mappingSkills)
                            columnCtr += 1
                            output(rowCtr, columnCtr) = GetRawSkillString(runningField.Value.NewFileData, mappingSkills)
                            columnCtr += 1
                            output(rowCtr, columnCtr) = Nothing
                        End If
                    Next
                    columnCtr += 1
                    output(rowCtr, columnCtr) = If(manualUpdate, "Y", "N")
                    columnCtr += 1
                    output(rowCtr, columnCtr) = If(systemUpdate, "Y", "N")
                    columnCtr += 1
                    output(rowCtr, columnCtr) = If(manualUpdate Or systemUpdate, "Y", "N")
                End If
            Next
            OnHeartbeatSub("")

            _cts.Token.ThrowIfCancellationRequested()
            OnHeartbeat("Opening Excel to write data")
            Dim outputFilename As String = Path.Combine(Path.GetDirectoryName(_newFilePath), String.Format("Output {0}.xlsx", Now.ToString("HH_mm_ss")))
            Using xlHlpr As New ExcelHelper(outputFilename, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                AddHandler xlHlpr.Heartbeat, AddressOf OnHeartbeat
                AddHandler xlHlpr.WaitingFor, AddressOf OnWaitingFor

                OnHeartbeat("Writing data to excel")
                Dim range As String = xlHlpr.GetNamedRange(1, output.GetLength(0) - 1, 1, output.GetLength(1) - 1)
                _cts.Token.ThrowIfCancellationRequested()
                xlHlpr.WriteArrayToExcel(output, range)
                OnHeartbeat(String.Format("Output file available at:{0}", Path.GetDirectoryName(outputFilename)))

                For i As Integer = 2 To 13
                    xlHlpr.AutoSizeColumnWidth(i)
                Next

                RemoveHandler xlHlpr.Heartbeat, AddressOf OnHeartbeat
                RemoveHandler xlHlpr.WaitingFor, AddressOf OnWaitingFor
            End Using
        End If
    End Function

    Private Sub CompareAllScores(ByRef scoreData As ScoreDetails, ByVal mappingList As List(Of String))
        If scoreData IsNot Nothing AndAlso scoreData.Fields IsNot Nothing AndAlso scoreData.Fields.Count > 0 Then
            For Each runningField In scoreData.Fields
                _cts.Token.ThrowIfCancellationRequested()
                CompareFieldScore(runningField.Value, runningField.Key, mappingList)
            Next
        End If
    End Sub

    Private Sub CompareFieldScore(ByRef fieldData As Field, ByVal fieldName As String, ByVal mappingList As List(Of String))
        If fieldData IsNot Nothing AndAlso (fieldData.OldFileValue IsNot Nothing OrElse fieldData.NewFileValue IsNot Nothing) Then
            Dim oldFileData() As String = fieldData.OldFileData
            Dim newFileData() As String = fieldData.NewFileData
            'Removed Skills
            _cts.Token.ThrowIfCancellationRequested()
            Dim removedSkills As List(Of String) = Nothing
            If oldFileData IsNot Nothing AndAlso oldFileData.Count > 0 AndAlso newFileData IsNot Nothing AndAlso newFileData.Count > 0 Then
                Dim indexNumber As Integer = 0
                For i As Integer = 0 To oldFileData.Count - 1
                    _cts.Token.ThrowIfCancellationRequested()
                    If Not IsSkillNameExists(oldFileData(i), newFileData, indexNumber) Then
                        If removedSkills Is Nothing Then removedSkills = New List(Of String)
                        removedSkills.Add(oldFileData(i))
                    Else
                        If fieldName.ToUpper <> "FIELD1" Then
                            If Not newFileData.Contains(oldFileData(i)) Then
                                If removedSkills Is Nothing Then removedSkills = New List(Of String)
                                removedSkills.Add(oldFileData(i))
                            End If
                        End If
                    End If
                Next
            ElseIf oldFileData IsNot Nothing AndAlso oldFileData.Count > 0 AndAlso (newFileData Is Nothing OrElse newFileData.Count = 0) Then
                For i As Integer = 0 To oldFileData.Count - 1
                    _cts.Token.ThrowIfCancellationRequested()
                    If removedSkills Is Nothing Then removedSkills = New List(Of String)
                    removedSkills.Add(oldFileData(i))
                Next
            End If
            'Added Skills
            _cts.Token.ThrowIfCancellationRequested()
            Dim addedSkills As List(Of String) = Nothing
            If oldFileData IsNot Nothing AndAlso oldFileData.Count > 0 AndAlso newFileData IsNot Nothing AndAlso newFileData.Count > 0 Then
                Dim indexNumber As Integer = 0
                For i As Integer = 0 To newFileData.Count - 1
                    _cts.Token.ThrowIfCancellationRequested()
                    If Not IsSkillNameExists(newFileData(i), oldFileData, indexNumber) Then
                        If addedSkills Is Nothing Then addedSkills = New List(Of String)
                        addedSkills.Add(newFileData(i))
                    Else
                        If fieldName.ToUpper <> "FIELD1" Then
                            If Not oldFileData.Contains(newFileData(i)) Then
                                If addedSkills Is Nothing Then addedSkills = New List(Of String)
                                addedSkills.Add(newFileData(i))
                            End If
                        End If
                    End If
                Next
            ElseIf newFileData IsNot Nothing AndAlso newFileData.Count > 0 AndAlso (oldFileData Is Nothing OrElse oldFileData.Count = 0) Then
                For i As Integer = 0 To newFileData.Count - 1
                    _cts.Token.ThrowIfCancellationRequested()
                    If addedSkills Is Nothing Then addedSkills = New List(Of String)
                    addedSkills.Add(newFileData(i))
                Next
            End If
            'Updated Skills
            _cts.Token.ThrowIfCancellationRequested()
            Dim updatedSkills As List(Of String) = Nothing
            If oldFileData IsNot Nothing AndAlso oldFileData.Count > 0 AndAlso newFileData IsNot Nothing AndAlso newFileData.Count > 0 Then
                If fieldName.ToUpper = "FIELD1" Then
                    For i As Integer = 0 To newFileData.Count - 1
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim indexNumber As Integer = 0
                        If IsSkillNameExists(newFileData(i), oldFileData, indexNumber) Then
                            If Not newFileData(i).Trim.ToUpper = oldFileData(indexNumber).Trim.ToUpper AndAlso Not oldFileData.Contains(newFileData(i)) Then
                                Dim scoreUpdate As String = GetScoreUpdate(oldFileData(indexNumber), newFileData(i), mappingList)
                                If scoreUpdate IsNot Nothing Then
                                    If updatedSkills Is Nothing Then updatedSkills = New List(Of String)
                                    updatedSkills.Add(scoreUpdate)
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            fieldData.Changes = New Difference With {
                .Added = addedSkills,
                .Removed = removedSkills,
                .Updated = updatedSkills,
                .OverallChange = GetSkillString(addedSkills, removedSkills, updatedSkills, mappingList)
            }
        End If
    End Sub

    Private Function GetSkillString(ByVal addedSkillList As List(Of String), ByVal removedSkillList As List(Of String), ByVal updatedSkillList As List(Of String), ByVal mappingList As List(Of String)) As String
        Dim ret As String = Nothing
        Dim skill As String = Nothing
        If addedSkillList IsNot Nothing AndAlso addedSkillList.Count > 0 Then
            For Each runningSkill In addedSkillList
                _cts.Token.ThrowIfCancellationRequested()
                If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(runningSkill).Trim, StringComparer.OrdinalIgnoreCase) Then
                    skill = String.Format("{0}{1}(+){2}", skill, vbNewLine, runningSkill)
                Else
                    skill = String.Format("{0}{1}(+){2}**************", skill, vbNewLine, runningSkill)
                End If
            Next
        End If
        If removedSkillList IsNot Nothing AndAlso removedSkillList.Count > 0 Then
            For Each runningSkill In removedSkillList
                _cts.Token.ThrowIfCancellationRequested()
                If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(runningSkill).Trim, StringComparer.OrdinalIgnoreCase) Then
                    skill = String.Format("{0}{1}(-){2}", skill, vbNewLine, runningSkill)
                Else
                    skill = String.Format("{0}{1}(-){2}**************", skill, vbNewLine, runningSkill)
                End If
            Next
        End If
        If updatedSkillList IsNot Nothing AndAlso updatedSkillList.Count > 0 Then
            For Each runningSkill In updatedSkillList
                _cts.Token.ThrowIfCancellationRequested()
                skill = String.Format("{0}{1}{2}", skill, vbNewLine, runningSkill)
            Next
        End If
        If skill IsNot Nothing AndAlso skill.Count > 0 Then
            ret = skill.Trim
        End If
        Return ret
    End Function

    Private Function GetRawSkillString(ByVal skillList() As String, ByVal mappingList As List(Of String)) As String
        Dim ret As String = Nothing
        Dim skill As String = Nothing
        If skillList IsNot Nothing AndAlso skillList.Count > 0 Then
            For Each runningSkill In skillList
                _cts.Token.ThrowIfCancellationRequested()
                If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(runningSkill).Trim, StringComparer.OrdinalIgnoreCase) Then
                    skill = String.Format("{0}{1}{2}", skill, vbNewLine, runningSkill)
                Else
                    skill = String.Format("{0}{1}{2}**************", skill, vbNewLine, runningSkill)
                End If
            Next
        End If
        If skill IsNot Nothing AndAlso skill.Count > 0 Then
            ret = skill.Trim
        End If
        Return ret
    End Function

    Private Function GetScoreUpdate(ByVal oldSkillScore As String, ByVal newSkillScore As String, ByVal mappingList As List(Of String)) As String
        Dim ret As String = Nothing
        Dim oldScore As String = GetScore(oldSkillScore)
        Dim newScore As String = GetScore(newSkillScore)
        _cts.Token.ThrowIfCancellationRequested()
        If oldScore IsNot Nothing AndAlso newScore IsNot Nothing Then
            If IsNumeric(oldScore) AndAlso IsNumeric(newScore) Then
                Dim delta As Decimal = Val(newScore) - Val(oldScore)
                If delta > 0 Then
                    If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(newSkillScore).Trim, StringComparer.OrdinalIgnoreCase) Then
                        ret = String.Format("(+){0}({1})", GetSkillName(newSkillScore), delta)
                    Else
                        ret = String.Format("(+){0}({1})**************", GetSkillName(newSkillScore), delta)
                    End If
                ElseIf delta < 0 Then
                    If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(newSkillScore).Trim, StringComparer.OrdinalIgnoreCase) Then
                        ret = String.Format("(-){0}({1})", GetSkillName(newSkillScore), Math.Abs(delta))
                    Else
                        ret = String.Format("(-){0}({1})**************", GetSkillName(newSkillScore), Math.Abs(delta))
                    End If
                End If
            Else
                Dim oldSubScore As String = oldScore(1)
                Dim newSubScore As String = newScore(1)
                If Val(newSubScore) > Val(oldSubScore) Then
                    If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(newSkillScore).Trim, StringComparer.OrdinalIgnoreCase) Then
                        ret = String.Format("(+){0}", newSkillScore)
                    Else
                        ret = String.Format("(+){0}**************", newSkillScore)
                    End If
                ElseIf Val(newSubScore) < Val(oldSubScore) Then
                    If mappingList IsNot Nothing AndAlso mappingList.Contains(GetSkillName(newSkillScore).Trim, StringComparer.OrdinalIgnoreCase) Then
                        ret = String.Format("(-){0}", newSkillScore)
                    Else
                        ret = String.Format("(-){0}**************", newSkillScore)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetScore(ByVal skillScore As String) As String
        Dim ret As String = ""
        If skillScore IsNot Nothing AndAlso skillScore.Contains("(") AndAlso skillScore.Contains(")") Then
            Dim firstIndex As Integer = skillScore.IndexOf("(")
            Dim secondIndex As Integer = skillScore.IndexOf(")")
            ret = skillScore.Substring(firstIndex + 1, secondIndex - (firstIndex + 1))
        End If
        Return ret
    End Function

    Private Function GetSkillName(ByVal skillScore As String) As String
        Dim ret As String = ""
        If skillScore IsNot Nothing AndAlso skillScore.Contains("(") AndAlso skillScore.Contains(")") Then
            ret = skillScore.Substring(0, skillScore.IndexOf("("))
        End If
        Return ret
    End Function

    Private Function IsSkillNameExists(ByVal skillNameWithScore As String, ByVal skillDataArray() As String, ByRef indexNumber As Integer) As Boolean
        Dim ret As Boolean = False
        If skillDataArray IsNot Nothing AndAlso skillDataArray.Count > 0 Then
            Dim skillName As String = GetSkillName(skillNameWithScore)
            For i As Integer = 0 To skillDataArray.Count - 1
                _cts.Token.ThrowIfCancellationRequested()
                Dim runningSkillName As String = GetSkillName(skillDataArray(i))
                If runningSkillName.Trim.ToUpper = skillName.Trim.ToUpper Then
                    ret = True
                    indexNumber = i
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Sub ReadRequiedDataFromFile(ByVal fileName As String, ByVal fileType As String, ByRef empScoreDetails As Dictionary(Of String, ScoreDetails))
        OnHeartbeat(String.Format("Opening {0} File", fileType))
        Using xlHlpr As New ExcelHelper(fileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            AddHandler xlHlpr.Heartbeat, AddressOf OnHeartbeat
            AddHandler xlHlpr.WaitingFor, AddressOf OnWaitingFor

            Dim allSheets As List(Of String) = xlHlpr.GetExcelSheetsName()
            If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                Dim dataSheet As String = Nothing
                For Each runningSheet In allSheets
                    _cts.Token.ThrowIfCancellationRequested()
                    If runningSheet.Contains("BFSI") Then
                        dataSheet = runningSheet
                        Exit For
                    End If
                Next
                If dataSheet IsNot Nothing Then
                    xlHlpr.SetActiveSheet(dataSheet)
                    OnHeartbeat(String.Format("Checking schema of {0} File", fileType))
                    xlHlpr.CheckExcelSchema(_fileSchema.Values.ToArray)
                    xlHlpr.UnFilterSheet(dataSheet)

                    OnHeartbeat(String.Format("Reading data from {0} File", fileType))
                    Dim scoreData As Object(,) = xlHlpr.GetExcelInMemory()
                    If scoreData IsNot Nothing Then
                        OnHeartbeatSub(String.Format("Seraching required column numbers from {0} File", fileType))
                        Dim empColumnNumber As Integer = GetColumnOf2DArray(scoreData, 1, _fileSchema("Emp No"))
                        Dim dataColumnNumbers As Dictionary(Of String, Integer) = Nothing
                        For Each runningColumn In _fileSchema
                            _cts.Token.ThrowIfCancellationRequested()
                            If Not runningColumn.Key = "Emp No" Then
                                Dim column As Integer = GetColumnOf2DArray(scoreData, 1, runningColumn.Value)

                                If dataColumnNumbers Is Nothing Then dataColumnNumbers = New Dictionary(Of String, Integer)
                                dataColumnNumbers.Add(runningColumn.Key, column)
                            End If
                        Next
                        If dataColumnNumbers IsNot Nothing AndAlso dataColumnNumbers.Count > 0 Then
                            If empScoreDetails Is Nothing Then empScoreDetails = New Dictionary(Of String, ScoreDetails)
                            For rowCounter As Integer = 2 To scoreData.GetLength(0) - 1
                                _cts.Token.ThrowIfCancellationRequested()
                                OnHeartbeatSub(String.Format("Reading required columns data # {0}/{1}", rowCounter - 1, scoreData.GetLength(0) - 2))
                                Dim empID As String = scoreData(rowCounter, empColumnNumber)
                                If empID IsNot Nothing AndAlso empID <> "" Then
                                    Dim fields As Dictionary(Of String, Field) = Nothing
                                    If empScoreDetails.ContainsKey(empID) Then
                                        fields = empScoreDetails(empID).Fields
                                    Else
                                        fields = New Dictionary(Of String, Field)
                                        empScoreDetails.Add(empID, New ScoreDetails With {.EmpNo = empID, .Fields = fields})
                                    End If
                                    For Each runningDataColumn In dataColumnNumbers
                                        _cts.Token.ThrowIfCancellationRequested()
                                        Dim runningField As Field = Nothing
                                        If fields IsNot Nothing AndAlso fields.Count > 0 AndAlso fields.ContainsKey(runningDataColumn.Key) Then
                                            runningField = fields(runningDataColumn.Key)
                                        Else
                                            runningField = New Field
                                            runningField.FieldName = _fileSchema(runningDataColumn.Key)
                                            fields.Add(runningDataColumn.Key, runningField)
                                        End If
                                        Dim fieldValue As String = scoreData(rowCounter, runningDataColumn.Value)
                                        If fileType.ToUpper = "OLD" Then
                                            runningField.OldFileValue = fieldValue
                                        ElseIf fileType.ToUpper = "NEW" Then
                                            runningField.NewFileValue = fieldValue
                                        End If
                                    Next
                                End If
                            Next
                            OnHeartbeatSub("")
                        End If
                    End If
                Else
                    Throw New ApplicationException(String.Format("BFSI sheet not found in {0} File", fileType))
                End If
            End If

            RemoveHandler xlHlpr.Heartbeat, AddressOf OnHeartbeat
            RemoveHandler xlHlpr.WaitingFor, AddressOf OnWaitingFor

            OnHeartbeat(String.Format("Closing {0} File", fileType))
        End Using
    End Sub

    Private Function GetColumnOf2DArray(ByVal array As Object(,), ByVal rowNumber As Integer, ByVal searchData As String) As Integer
        Dim ret As Integer = Integer.MinValue
        If array IsNot Nothing AndAlso searchData IsNot Nothing Then
            For column As Integer = 1 To array.GetLength(1)
                _cts.Token.ThrowIfCancellationRequested()
                If array(rowNumber, column) IsNot Nothing AndAlso
                    array(rowNumber, column).ToString.ToUpper = searchData.ToUpper Then
                    ret = column
                    If ret <> Integer.MinValue Then Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Function ReadRequiedMappingFile(ByVal fileName As String) As List(Of String)
        Dim ret As List(Of String) = Nothing
        OnHeartbeat(String.Format("Opening Mapping File"))
        Using xl As New ExcelHelper(fileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            AddHandler xl.Heartbeat, AddressOf OnHeartbeat
            AddHandler xl.WaitingFor, AddressOf OnWaitingFor

            Dim allSheets As List(Of String) = xl.GetExcelSheetsName()
            If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                Dim sheetCounter As Integer = 0
                For Each runningSheet In allSheets
                    _cts.Token.ThrowIfCancellationRequested()
                    sheetCounter += 1
                    xl.SetActiveSheet(runningSheet)
                    _cts.Token.ThrowIfCancellationRequested()
                    xl.UnFilterSheet(runningSheet)
                    _cts.Token.ThrowIfCancellationRequested()
                    OnHeartbeat(String.Format("Checking schema of '{0}' sheet in {1}", runningSheet, fileName))
                    xl.CheckExcelSchema(_mappingFileSchema.Values.ToArray)
                    _cts.Token.ThrowIfCancellationRequested()
                    Dim skillRowColumnNumber As KeyValuePair(Of Integer, Integer) = xl.FindAll(_mappingFileSchema("WFT Practice"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault
                    Dim skillName As String = xl.GetData(skillRowColumnNumber.Key + 1, skillRowColumnNumber.Value)
                    If skillName IsNot Nothing AndAlso skillName <> "" Then
                        OnHeartbeatSub(String.Format("Reading mapping skills data # {0}/{1}", sheetCounter, allSheets.Count))
                        Dim subSkillColumnNumber As Integer = xl.FindAll(_mappingFileSchema("WFT Subskills"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                        Dim subSkillLastRow As Integer = xl.GetLastRow(subSkillColumnNumber)
                        For subskillRow As Integer = 2 To subSkillLastRow
                            _cts.Token.ThrowIfCancellationRequested()
                            Dim subskill As String = xl.GetData(subskillRow, subSkillColumnNumber)
                            If subskill IsNot Nothing AndAlso subskill <> "" Then
                                Dim nomenclatureLastColumn As Integer = xl.GetLastCol(subskillRow)
                                For nomenclatureCol As Integer = subSkillColumnNumber + 1 To nomenclatureLastColumn
                                    _cts.Token.ThrowIfCancellationRequested()
                                    Dim nomenclature As String = xl.GetData(subskillRow, nomenclatureCol)
                                    If nomenclature IsNot Nothing AndAlso nomenclature <> "" Then
                                        If ret Is Nothing Then ret = New List(Of String)
                                        ret.Add(nomenclature.Trim)
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
                OnHeartbeatSub("")
            End If

            RemoveHandler xl.Heartbeat, AddressOf OnHeartbeat
            RemoveHandler xl.WaitingFor, AddressOf OnWaitingFor

            OnHeartbeat(String.Format("Closing Mapping File"))
        End Using
        Return ret
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls
    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
