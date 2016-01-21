Imports NextGen.Framework.Managers.AppObjectMgr
Imports NextGen.Framework.Managers.OwnedObjectMgr
Imports NextGen.Framework.Core.XML.XMLHelpers
Imports NextGen.Framework.Managers.AppObjectMgr.Presentation
Imports NextGen.Framework.Managers.ConfigurationMgr
Imports NextGen.Framework.Managers.QueryMgr

Public Class UpdatePresentationsProcess
    Inherits AppDomainFor3E

    Public Sub New()
        Call MyBase.New()
    End Sub

    Public Sub Initialise()
        Call Me.DoCallBack(AddressOf Internal.Initialise)
    End Sub

    Public Sub UpdatePresentation(ByVal presentationAppObjectCode As String, ByVal runId As String)
        If Me.IsDisposed Then Throw New ObjectDisposedException("UpdatePresentationsProcess")

        Me.PresentationToUpdate = presentationAppObjectCode
        Me.OriginalRunId = runId
        Call Me.DoCallBack(AddressOf Internal.DoUpdate)
    End Sub

    ' ReSharper disable once ClassNeverInstantiated.Local
    Private Class Internal
        Public Shared Sub Initialise()
            Dim configMgr As New NxConfigMgr()
            NxConfigMgr.Initialize(configMgr)
            Dim queryMgrConfig As NxConfigSection = configMgr.GetConfigSection("nextGen/managers/globalManagers/queryManager")
            Call SharedQueryMgr.Initialize(queryMgrConfig)
        End Sub

        Public Shared Sub DoUpdate()
            Try
                Dim p As NextGen.Framework.Managers.AppObjectMgr.Presentation.Presentation = GetPresentation(PresentationToUpdate)

                Dim needsSavingForDistinctLayout As Boolean = False
                If Not p.ReportLayout.IsUseDistinctLayout.Value Then
                    p.ReportLayout.IsUseDistinctLayout.Value = True
                    Call Console.WriteLine("{0}: {1} distinct layout", PresentationToUpdate, If(InLiveMode, "Setting", "Would set"))
                    needsSavingForDistinctLayout = True
                End If

                Dim needsSavingForRunId As Boolean = False
                If Not String.IsNullOrWhiteSpace(OriginalRunId) Then
                    If TryUpdateRunId(p, OriginalRunId) Then
                        needsSavingForRunId = True
                    End If
                End If

                If Not needsSavingForDistinctLayout AndAlso Not needsSavingForRunId Then
                    Exit Sub
                End If

                If Not InLiveMode Then
                    Return
                End If

                Call NxOwnedObjectMgr.SaveApplicationObject(p, p.ID.Value, Environment.UserName, "Update Presentations", True)
                Call Console.WriteLine("{0}: Updated", PresentationToUpdate)
            Catch ex As Exception
                ExceptionToPassBack = ex
            End Try
        End Sub

        Private Shared Function TryUpdateRunId(ByVal presentation As NextGen.Framework.Managers.AppObjectMgr.Presentation.Presentation, ByVal runId As String) As Boolean
            Dim pageDimensions As Dimensions = presentation.PAGEDIMENSIONS
            If pageDimensions Is Nothing Then
                pageDimensions = New Dimensions(presentation)
                presentation.PAGEDIMENSIONS = pageDimensions
            End If

            Dim runDimension As RunDimension = pageDimensions.OfType(Of RunDimension)().SingleOrDefault()
            If runDimension Is Nothing Then
                runDimension = New RunDimension(presentation.PAGEDIMENSIONS)
                If pageDimensions.Count = 0 Then
                    Call pageDimensions.Add(runDimension)
                Else
                    Call pageDimensions.Insert(0, runDimension)
                End If
            End If

            Dim run As Run = runDimension.OfType(Of Run)().SingleOrDefault()
            Dim runExists As Boolean = run IsNot Nothing
            If Not runExists Then
                run = New Run(runDimension)
            End If

            If String.Equals(runId, run.ID.Value) Then Return False

            ' must set id before adding to runDimension
            Call Console.WriteLine("{0}: {1} run id from {2} to {3}", presentation.ID.Value, If(InLiveMode, "Setting", "Would set"), If(run.ID.Value, "(null)"), runId)
            run.ID.Value = runId

            If Not runExists Then
                Call runDimension.Add(run)
            End If

            Return True
        End Function

        Private Shared Function GetPresentation(ByVal aoc As String) As Presentation
            Dim result As Presentation = CType(GetAppObject(aoc, NxDocumentType.Presentation), Presentation)
            Return result
        End Function

        Private Shared Function GetAppObject(ByVal aoc As String, ByVal aot As NxDocumentType) As IAppObject
            Return NxAppObjectMgr.GetAppObject(aoc, aot, True, AddressOf NxOwnedObjectMgr.AppObjectLoader)
        End Function

        Private Shared ReadOnly Property PresentationToUpdate() As String
            Get
                Return GetDataFromCurrentDomain(Of String)()
            End Get
        End Property

        Private Shared ReadOnly Property OriginalRunId() As String
            Get
                Return GetDataFromCurrentDomain(Of String)()
            End Get
        End Property
    End Class

    Private WriteOnly Property PresentationToUpdate As String
        Set(value As String)
            Call SetDomainData(Of String)(data:=value)
        End Set
    End Property

    Private WriteOnly Property OriginalRunId As String
        Set(value As String)
            Call SetDomainData(Of String)(data:=value)
        End Set
    End Property
End Class
