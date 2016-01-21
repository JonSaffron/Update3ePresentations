Imports System.IO

Module Program
    Public InLiveMode As Boolean

    Function Main(ByVal args As String()) As Integer
        Try
            Call DisplayHeader(args)

            Dim listOfPresentations As PresentationDefinition() = GetAllCustomPresentationsAndTheirOrigins().Where(Function(item) item.OriginalAppObjectType = "ReportLayout").ToArray()

            Call GetLastRunIdsForPresentations(listOfPresentations)

            Call UpdatePresentations(listOfPresentations)

            Return 0
        Catch ex As Exception
            Call Console.WriteLine()
            Call Console.WriteLine(ex.ToString())

            Return 1
        Finally
            If IsRunningInVisualStudio() Then
                Call Console.WriteLine()
                Call Console.WriteLine("Press Return to continue.")
                Call Console.ReadLine()
            End If
        End Try
    End Function

    Private Sub DisplayHeader(ByVal args As IEnumerable(Of String))
        Call Console.WriteLine("Update presentations in 3e")
        Call Console.WriteLine()

        Dim iterator As IEnumerator(Of String) = args.GetEnumerator()
        Do While iterator.MoveNext()
            Dim currentArg As String = iterator.Current
            If currentArg.Equals("RUN", StringComparison.Ordinal) Then
                InLiveMode = True
            End If
        Loop
    End Sub

    Private Function GetAllCustomPresentationsAndTheirOrigins() As PresentationDefinition()
        Call Console.WriteLine("Loading presentation information...")
        Using x As New GetCustomPresentationsProcess()
            Dim result As PresentationDefinition() = x.GetCustomPresentations()
            Call Console.WriteLine("Complete.")
            Call Console.WriteLine()
            Return result
        End Using
    End Function

    Private Sub GetLastRunIdsForPresentations(ByVal listOfPresentations As PresentationDefinition())
        Call Console.WriteLine("Getting run ids from 2.5...")
        For Each pd As PresentationDefinition In listOfPresentations
            pd.OriginalRunId = GetRunIdFromReportLayout(pd.OriginalAppObjectCode)
        Next
        Call Console.WriteLine()
    End Sub

    Private Sub UpdatePresentations(ByVal listOfPresentations As PresentationDefinition())
        Call Console.WriteLine("Updating presentations...")
        Using x As New UpdatePresentationsProcess
            Call x.Initialise()

            For Each pd As PresentationDefinition In listOfPresentations
                Call x.UpdatePresentation(pd.PresentationAppObjectCode, pd.OriginalRunId)
            Next
        End Using
        Call Console.WriteLine()
    End Sub

    Private Function IsRunningInVisualStudio() As Boolean
        If Process.GetCurrentProcess().MainModule.ModuleName.EndsWith(".vshost.exe", StringComparison.OrdinalIgnoreCase) Then
            Dim basePath As String = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
            If basePath.EndsWith("\bin\debug", StringComparison.OrdinalIgnoreCase) OrElse basePath.EndsWith("\bin\release", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        End If

        Return False
    End Function
End Module
