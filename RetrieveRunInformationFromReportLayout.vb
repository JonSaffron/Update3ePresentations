Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml

Public Module RetrieveRunInformationFromReportLayout
    Public Function GetRunIdFromReportLayout(ByVal reportLayoutAppObjectCode As String) As String
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("OldLive").ConnectionString)
            Call conn.Open()

            Using command As SqlCommand = BuildSqlCommand(reportLayoutAppObjectCode, conn)
                Using reader As SqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Call Console.WriteLine("{0}: Couldn't find definition", reportLayoutAppObjectCode)
                        Return Nothing
                    End If

                    Dim xml As String = reader.GetString(0)
                    Dim result As String = GetRunIdFromXml(xml, reportLayoutAppObjectCode)

                    If reader.Read() Then
                        Call Console.WriteLine("{0}: Unexpected multiple records encountered", reportLayoutAppObjectCode)
                        Return Nothing
                    End If

                    Return result
                End Using
            End Using
        End Using
    End Function

    Private Function BuildSqlCommand(ByVal reportLayoutAppObjectCode As String, ByVal connection As SqlConnection) As SqlCommand
        Dim sql As String =
            "SELECT		MetaXml " & _
            "FROM		NxFwkAppObjectData INNER JOIN NxFwkAppObject ON NxFwkAppObjectData.AppObjectId = NxFWKAppObject.NxFWKAppObjectID " & _
            "WHERE		NxFWKAppObject.AppObjectCode = @AppObjectCode " & _
            "AND			NxFWKAppObject.AppObjectTypeId = (SELECT NxFWKAppObjectTypeID FROM NxFWKAppObjectType WHERE NxFWKAppObjectType.AppObjectTypeCode = 'ReportLayout') " & _
            "AND			NxFWKAppObjectData.IsCustom = 1"
        Dim result As New SqlCommand(sql, connection)
        Dim parameter As SqlParameter = result.Parameters.Add("@AppObjectCode", SqlDbType.NVarChar)
        parameter.Value = reportLayoutAppObjectCode
        Return result
    End Function

    Private Function GetRunIdFromXml(xml As String, ByVal reportLayoutAppObjectCode As String) As String
        Dim xmlDoc As New XmlDocument()
        Call xmlDoc.LoadXml(xml)

        Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
        Call nsmgr.AddNamespace("ns", "http://elite.com/schemas/reportlayout")
        Dim runAttribute As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("@RunID", nsmgr)
        If runAttribute Is Nothing Then
            Call Console.WriteLine("{0}: No run id set.", reportLayoutAppObjectCode)
            Return Nothing
        End If

        Dim result As String = runAttribute.Value
        If String.IsNullOrWhiteSpace(result) Then
            Call Console.WriteLine("{0}: No run id set.", reportLayoutAppObjectCode)
            Return Nothing
        End If

        Return result
    End Function
End Module
