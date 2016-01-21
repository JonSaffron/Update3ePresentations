Imports NextGen.Framework.Managers.QueryMgr
Imports NextGen.Framework.OQL
Imports NextGen.Application

Public Class GetCustomPresentationsProcess
    Inherits AppDomainFor3E

    Public Sub New()
        Call MyBase.New()
    End Sub

    Public Function GetCustomPresentations() As PresentationDefinition()
        If Me.IsDisposed Then Throw New ObjectDisposedException("GetCustomPresentationsProcess")

        Call Me.DoCallBack(AddressOf Internal.GetListOfPresentations)
        Dim result As PresentationDefinition() = Presentations
        Return result
    End Function

    ' ReSharper disable once ClassNeverInstantiated.Local
    Private Class Internal
        Public Shared Sub GetListOfPresentations()
            Try
                Dim qs As ISession = GetQuerySession()
                Dim listOfPresentations As New List(Of PresentationDefinition)()

                Dim selectQ As ISelect = GetSelectForCustomPresentations()
                Using cursor As ICursor = qs.OpenCursor(selectQ)
                    While cursor.MoveNext()
                        Dim presentationAppObjectCode As String = cursor.GetString(0)
                        Dim item As PresentationDefinition = GetDefinition(qs, presentationAppObjectCode)
                        If item IsNot Nothing Then
                            Call listOfPresentations.Add(item)
                        End If
                    End While
                End Using

                Presentations = listOfPresentations.ToArray()

            Catch ex As Exception
                ExceptionToPassBack = ex
            End Try
        End Sub

        Private Shared Function GetSelectForCustomPresentations() As ISelect
            Dim tNxFwkAppObjectData As Query.NxFWKAppObjectData = Query.NxFWKAppObjectData.Current

            Dim selectQ As ISelect = QueryBuilder.SelectFrom(tNxFwkAppObjectData)
            Call selectQ.Values(tNxFwkAppObjectData.NxFWKAppObject.AppObjectCode)
            Call selectQ.Where(tNxFwkAppObjectData.NxFWKAppObject.NxFWKAppObjectType.AppObjectTypeCode.IsEqualTo("Presentation"))
            Call selectQ.Where(tNxFwkAppObjectData.IsCustom.IsTrue())
            Call selectQ.OrderBy(tNxFwkAppObjectData.NxFWKAppObject.AppObjectCode)
            Return selectQ
        End Function

        Private Shared Function GetDefinition(ByVal qs As ISession, ByVal presentationAppObjectCode As String) As PresentationDefinition
            Dim originalAppObjectType As String = Nothing
            Dim originalAppObjectCode As String = Nothing
            If Not TryGetOriginalAppObjectTypeAndCode(qs, presentationAppObjectCode, originalAppObjectType, originalAppObjectCode) Then
                Return Nothing
            End If

            Dim result As New PresentationDefinition()
            result.PresentationAppObjectCode = presentationAppObjectCode
            result.OriginalAppObjectType = originalAppObjectType
            result.OriginalAppObjectCode = originalAppObjectCode
            Return result
        End Function

        Private Shared Function TryGetOriginalAppObjectTypeAndCode(ByVal qs As ISession, ByVal presentationAppObjectCode As String, ByRef originalAppObjectType As String, ByRef originalAppObjectCode As String) As Boolean
            Dim selectQ As ISelect = GetSelectForOriginalTypeAndCode(presentationAppObjectCode)
            Using cursor As ICursor = qs.OpenCursor(selectQ)
                If Not cursor.MoveNext() Then
                    Call Console.WriteLine("{0}: No information about the original type and code", presentationAppObjectCode)
                    Return False
                End If

                originalAppObjectType = cursor.GetString(0)
                originalAppObjectCode = cursor.GetString(1)

                If cursor.MoveNext() Then
                    Call Console.WriteLine("{0}: Multiple alternative information available", presentationAppObjectCode)
                    Return False
                End If
            End Using

            Return True
        End Function

        Private Shared Function GetSelectForOriginalTypeAndCode(ByVal presentationAppObjectCode As String) As ISelect
            Dim tPresConv As Query.NxPresConv = Query.NxPresConv.Current

            Dim selectQ As ISelect = QueryBuilder.SelectFrom(tPresConv)
            Call selectQ.Distinct()
            Call selectQ.Values(tPresConv.OriginalAppObjectType, tPresConv.OriginalAppObjectCode)
            Call selectQ.Where(tPresConv.PresentationAppObjectCode.IsEqualTo(presentationAppObjectCode))
            Return selectQ
        End Function

        Private Shared Function GetQuerySession() As ISession
            Dim result As ISession = Utilities3ECommon.Database.GetQuerySession()
            Return result
        End Function

        Private Shared WriteOnly Property Presentations As PresentationDefinition()
            Set(value As PresentationDefinition())
                Call SetDataForCurrentDomain(Of PresentationDefinition())(data:=value)
            End Set
        End Property
    End Class

    Private ReadOnly Property Presentations As PresentationDefinition()
        Get
            Return GetDomainData(Of PresentationDefinition())()
        End Get
    End Property
End Class
