Public Class ScoreDetails
    Public Property EmpNo As String
    Public Property Fields As Dictionary(Of String, Field)
End Class

Public Class Field
    Public FieldName As String
    Public Property OldFileValue As String
    Public Property NewFileValue As String

    Private _OldFileData() As String
    Public ReadOnly Property OldFileData As String()
        Get
            If _OldFileData Is Nothing OrElse _OldFileData.Count = 0 Then
                If Me.OldFileValue IsNot Nothing Then
                    Dim seperator As String() = {"),"}
                    _OldFileData = Me.OldFileValue.Split(seperator, StringSplitOptions.RemoveEmptyEntries)
                    If _OldFileData IsNot Nothing AndAlso _OldFileData.Count > 0 Then
                        For data As Integer = 0 To _OldFileData.Count - 2
                            _OldFileData(data) = String.Format("{0})", _OldFileData(data))
                        Next
                    End If
                End If
            End If
            Return _OldFileData
        End Get
    End Property

    Private _NewFileData() As String
    Public ReadOnly Property NewFileData As String()
        Get
            If _NewFileData Is Nothing OrElse _NewFileData.Count = 0 Then
                If Me.NewFileValue IsNot Nothing Then
                    Dim seperator As String() = {"),"}
                    _NewFileData = Me.NewFileValue.Split(seperator, StringSplitOptions.RemoveEmptyEntries)
                    If _NewFileData IsNot Nothing AndAlso _NewFileData.Count > 0 Then
                        For data As Integer = 0 To _NewFileData.Count - 2
                            _NewFileData(data) = String.Format("{0})", _NewFileData(data))
                        Next
                    End If
                End If
            End If
            Return _NewFileData
        End Get
    End Property

    Public Property Changes As Difference
End Class

Public Class Difference
    Public Property Added As List(Of String)
    Public Property Removed As List(Of String)
    Public Property Updated As List(Of String)
    Public Property OverallChange As String
End Class

Public Class MappingDetails
    Public Property Nomenclature As String
    Public Property SubSkill As String
    Public Property SkillLevel As String
    Public Property Practice As String
End Class