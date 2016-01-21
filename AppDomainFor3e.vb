Imports System.Configuration
Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Security
Imports System.Security.Permissions
Imports System.Security.Policy

<Serializable> Public MustInherit Class AppDomainFor3E
    Implements IDisposable

    Private _appDomain As AppDomain
    Private _copyOfAssembly As String

    Protected Sub New()
        Me._appDomain = CreateAppDomain()
    End Sub

    Private Function CreateAppDomain() As AppDomain
        Dim rootPath As String = ConfigurationManager.AppSettings("RootPath")
        Dim frameworkStaging As String = ConfigurationManager.AppSettings("FrameworkStaging")
        Dim pathToUtilities3EDirectory As String = Path.Combine(rootPath, "Utilities\Utilities3E")
        Dim pathToStagingDirectory As String = Path.Combine(rootPath, frameworkStaging)

        Dim ads As New AppDomainSetup()
        ads.ApplicationBase = rootPath
        ads.PrivateBinPath = String.Format("{0};{1}", pathToUtilities3EDirectory, pathToStagingDirectory)
        ads.PrivateBinPathProbe = "exclude application base path"
        ads.LoaderOptimization = LoaderOptimization.SingleDomain
        ads.DisallowBindingRedirects = True
        ads.ConfigurationFile = Path.Combine(pathToUtilities3EDirectory, "Utilities3e.exe.config")

        Dim securityEvidence As New Evidence()
        Call securityEvidence.AddHostEvidence(New Zone(SecurityZone.MyComputer))

        Dim ps As New PermissionSet(PermissionState.Unrestricted)

        Dim result As AppDomain = AppDomain.CreateDomain("CompileAndRun", securityEvidence, ads, ps)
        Me._copyOfAssembly = Path.Combine(pathToUtilities3EDirectory, Assembly.GetExecutingAssembly().GetName().Name) & ".dll"
        File.Copy(Assembly.GetExecutingAssembly().Location, Me._copyOfAssembly, True)

        Call result.SetData("PathToUtilities3EDirectory", pathToUtilities3EDirectory)
        Call result.SetData("InLiveMode", InLiveMode)
        Return result
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not IsDisposed Then
            Call AppDomain.Unload(Me._appDomain)
            Me._appDomain = Nothing
            Call File.Delete(Me._copyOfAssembly)
        End If
    End Sub

    Protected ReadOnly Property IsDisposed As Boolean
        Get
            Dim result = Me._appDomain Is Nothing
            Return result
        End Get
    End Property

    Protected Sub DoCallBack(ByVal callBackDelegate As CrossAppDomainDelegate)
        Call ClearException()
        Call Me._appDomain.DoCallBack(callBackDelegate)
        Dim ex As Exception = Me.Exception
        If ex IsNot Nothing Then Throw ex
    End Sub

    Protected Function GetDomainData(Of T)(<CallerMemberName> Optional ByVal name As String = Nothing) As T
        Dim result As T = CType(Me._appDomain.GetData(name), T)
        Return result
    End Function

    Protected Sub SetDomainData(Of T)(<CallerMemberName> Optional ByVal name As String = Nothing, Optional ByVal data As T = Nothing)
        Call Me._appDomain.SetData(name, data)
    End Sub

    Protected Shared Function GetDataFromCurrentDomain(Of T)(<CallerMemberName> Optional ByVal name As String = Nothing) As T
        Dim result As T = CType(AppDomain.CurrentDomain.GetData(name), T)
        Return result
    End Function

    Protected Shared Sub SetDataForCurrentDomain(Of T)(<CallerMemberName> Optional ByVal name As String = Nothing, Optional ByVal data As T = Nothing)
        Call AppDomain.CurrentDomain.SetData(name, data)
    End Sub

    Private Sub ClearException()
        Call SetDomainData(Of Exception)("Exception", Nothing)
    End Sub

    Private ReadOnly Property Exception As Exception
        Get
            Return GetDomainData(Of Exception)()
        End Get
    End Property

    Protected Shared WriteOnly Property ExceptionToPassBack As Exception
        Set(value As Exception)
            Call SetDataForCurrentDomain(Of Exception)("Exception", value)
        End Set
    End Property

    Protected Shared ReadOnly Property InLiveMode As Boolean
        Get
            Return GetDataFromCurrentDomain(Of Boolean)()
        End Get
    End Property
End Class
