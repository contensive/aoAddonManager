
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports System.Reflection

Namespace Contensive.Addons.AddonManager.Models
    Public MustInherit Class baseModel
        '
        '====================================================================================================
        '-- const must be set in derived clases
        '
        'Public Const contentName As String = "" '<------ set content name
        'Public Const contentTableName As String = "" '<------ set to tablename for the primary content (used for cache names)
        'Public Const contentDataSource As String = "" '<----- set to datasource if not default
        '
        '====================================================================================================
        ' -- instance properties
        Public Property id As Integer
        Public Property name As String
        Public Property ccguid As String
        Public Property Active As Boolean
        Public Property ContentControlID As Integer
        Public Property CreatedBy As Integer
        Public Property CreateKey As Integer
        Public Property DateAdded As Date
        Public Property ModifiedBy As Integer
        Public Property ModifiedDate As Date
        Public Property SortOrder As String
        '
        '====================================================================================================
        Private Shared Function derivedContentName(derivedType As Type) As String
            Dim fieldInfo As FieldInfo = derivedType.GetField("contentName")
            If (fieldInfo Is Nothing) Then
                Throw New ApplicationException("Class [" & derivedType.Name & "] must declare constant [contentName].")
            Else
                Return fieldInfo.GetRawConstantValue().ToString()
            End If
        End Function
        '
        '====================================================================================================
        Private Shared Function derivedContentTableName(derivedType As Type) As String
            Dim fieldInfo As FieldInfo = derivedType.GetField("contentTableName")
            If (fieldInfo Is Nothing) Then
                Throw New ApplicationException("Class [" & derivedType.Name & "] must declare constant [contentTableName].")
            Else
                Return fieldInfo.GetRawConstantValue().ToString()
            End If
        End Function
        '
        '====================================================================================================
        Private Shared Function contentDataSource(derivedType As Type) As String
            Dim fieldInfo As FieldInfo = derivedType.GetField("contentTableName")
            If (fieldInfo Is Nothing) Then
                Throw New ApplicationException("Class [" & derivedType.Name & "] must declare constant [contentTableName].")
            Else
                Return fieldInfo.GetRawConstantValue().ToString()
            End If
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' Create an empty object. needed for deserialization
        ''' </summary>
        Public Sub New()
            '
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' Add a new recod to the db and open it. Starting a new model with this method will use the default values in Contensive metadata (active, contentcontrolid, etc).
        ''' include callersCacheNameList to get a list of cacheNames used to assemble this response
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <returns></returns>
        Protected Shared Function add(Of T As baseModel)(cp As CPBaseClass) As T
            Dim result As T = Nothing
            Try
                Dim instanceType As Type = GetType(T)
                Dim contentName As String = derivedContentName(instanceType)
                result = create(Of T)(cp, cp.Content.AddRecord(contentName))
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' return a new model with the data selected. All cacheNames related to the object will be added to the cacheNameList.
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="recordId">The id of the record to be read into the new object</param>
        Protected Shared Function create(Of T As baseModel)(cp As CPBaseClass, recordId As Integer) As T
            Dim result As T = Nothing
            Try
                If recordId > 0 Then
                    Dim instanceType As Type = GetType(T)
                    Dim contentName As String = derivedContentName(instanceType)
                    Dim cs As CPCSBaseClass = cp.CSNew()
                    If cs.Open(contentName, "(id=" & recordId.ToString() & ")") Then
                        result = loadRecord(Of T)(cp, cs)
                    End If
                    cs.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' open an existing object
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="recordGuid"></param>
        Protected Shared Function create(Of T As baseModel)(cp As CPBaseClass, recordGuid As String) As T
            Dim result As T = Nothing
            Try
                Dim instanceType As Type = GetType(T)
                Dim contentName As String = derivedContentName(instanceType)
                Dim cs As CPCSBaseClass = cp.CSNew()
                If cs.Open(contentName, "(ccGuid=" & cp.Db.EncodeSQLText(recordGuid) & ")") Then
                    result = loadRecord(Of T)(cp, cs)
                End If
                cs.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' open an existing object
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="recordName"></param>
        Protected Shared Function createByName(Of T As baseModel)(cp As CPBaseClass, recordName As String) As T
            Dim result As T = Nothing
            Try
                If Not String.IsNullOrEmpty(recordName) Then
                    Dim instanceType As Type = GetType(T)
                    Dim contentName As String = derivedContentName(instanceType)
                    Dim cs As CPCSBaseClass = cp.CSNew()
                    If cs.Open(contentName, "(name=" & cp.Db.EncodeSQLText(recordName) & ")", "id") Then
                        result = loadRecord(Of T)(cp, cs)
                    End If
                    cs.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' open an existing object
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="cs"></param>
        Private Shared Function loadRecord(Of T As baseModel)(cp As CPBaseClass, cs As CPCSBaseClass) As T
            Dim instance As T = Nothing
            Try
                If cs.OK() Then
                    Dim instanceType As Type = GetType(T)
                    Dim tableName As String = derivedContentTableName(instanceType)
                    instance = DirectCast(Activator.CreateInstance(instanceType), T)
                    For Each resultProperty As PropertyInfo In instance.GetType().GetProperties(BindingFlags.Instance Or BindingFlags.Public)
                        Select Case resultProperty.Name.ToLower()
                            Case "specialcasefield"
                            Case "sortorder"
                                '
                                ' -- customization for pc, could have been in default property, db default, etc.
                                Dim sortOrder As String = cs.GetText(resultProperty.Name)
                                If (String.IsNullOrEmpty(sortOrder)) Then
                                    sortOrder = "9999"
                                End If
                                resultProperty.SetValue(instance, sortOrder, Nothing)
                            Case Else
                                Select Case resultProperty.PropertyType.Name
                                    Case "Int32"
                                        resultProperty.SetValue(instance, cs.GetInteger(resultProperty.Name), Nothing)
                                    Case "Boolean"
                                        resultProperty.SetValue(instance, cs.GetBoolean(resultProperty.Name), Nothing)
                                    Case "DateTime"
                                        resultProperty.SetValue(instance, cs.GetDate(resultProperty.Name), Nothing)
                                    Case "Double"
                                        resultProperty.SetValue(instance, cs.GetNumber(resultProperty.Name), Nothing)
                                    Case Else
                                        resultProperty.SetValue(instance, cs.GetText(resultProperty.Name), Nothing)
                                End Select
                        End Select
                    Next
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return instance
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' save the instance properties to a record with matching id. If id is not provided, a new record is created.
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <returns></returns>
        Protected Function save(cp As CPBaseClass) As Integer
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim instanceType As Type = Me.GetType()
                Dim contentName As String = derivedContentName(instanceType)
                Dim tableName As String = derivedContentTableName(instanceType)
                If (id > 0) Then
                    If Not cs.Open(contentName, "id=" & id) Then
                        Dim message As String = "Unable to open record in content [" & contentName & "], with id [" & id & "]"
                        cs.Close()
                        id = 0
                        Throw New ApplicationException(message)
                    End If
                Else
                    If Not cs.Insert(contentName) Then
                        cs.Close()
                        id = 0
                        Throw New ApplicationException("Unable to insert record in content [" & contentName & "]")
                    End If
                End If
                For Each resultProperty As PropertyInfo In Me.GetType().GetProperties(BindingFlags.Instance Or BindingFlags.Public)
                    Select Case resultProperty.Name.ToLower()
                        Case "id"
                            id = cs.GetInteger("id")
                        Case "ccguid"
                            If (String.IsNullOrEmpty(ccguid)) Then
                                ccguid = Guid.NewGuid().ToString()
                            End If
                            Dim value As String
                            value = resultProperty.GetValue(Me, Nothing).ToString()
                            cs.SetField(resultProperty.Name, value)
                        Case Else
                            Select Case resultProperty.PropertyType.Name
                                Case "Int32"
                                    Dim value As Integer
                                    Integer.TryParse(resultProperty.GetValue(Me, Nothing).ToString(), value)
                                    cs.SetField(resultProperty.Name, value.ToString())
                                Case "Boolean"
                                    Dim value As Boolean
                                    Boolean.TryParse(resultProperty.GetValue(Me, Nothing).ToString(), value)
                                    cs.SetField(resultProperty.Name, value.ToString())
                                Case "DateTime"
                                    Dim value As Date
                                    Date.TryParse(resultProperty.GetValue(Me, Nothing).ToString(), value)
                                    cs.SetField(resultProperty.Name, value.ToString())
                                Case "Double"
                                    Dim value As Double
                                    Double.TryParse(resultProperty.GetValue(Me, Nothing).ToString(), value)
                                    cs.SetField(resultProperty.Name, value.ToString())
                                Case Else
                                    Dim value As String
                                    value = resultProperty.GetValue(Me, Nothing).ToString()
                                    cs.SetField(resultProperty.Name, value)
                            End Select
                    End Select
                Next
                cs.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return id
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' delete an existing database record by id
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="recordId"></param>
        Protected Shared Sub delete(Of T As baseModel)(cp As CPBaseClass, recordId As Integer)
            Try
                If (recordId > 0) Then
                    Dim instanceType As Type = GetType(T)
                    Dim contentName As String = derivedContentName(instanceType)
                    Dim tableName As String = derivedContentTableName(instanceType)
                    cp.Content.Delete(contentName, "id=" & recordId.ToString)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' delete an existing database record by guid
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ccguid"></param>
        Protected Shared Sub delete(Of T As baseModel)(cp As CPBaseClass, ccguid As String)
            Try
                If (Not String.IsNullOrEmpty(ccguid)) Then
                    Dim instanceType As Type = GetType(T)
                    Dim contentName As String = derivedContentName(instanceType)
                    Dim instance As baseModel = create(Of baseModel)(cp, ccguid)
                    If (instance IsNot Nothing) Then
                        cp.Content.Delete(contentName, "(ccguid=" & cp.Db.EncodeSQLText(ccguid) & ")")
                    End If
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' pattern get a list of objects from this model
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="sqlCriteria"></param>
        ''' <returns></returns>
        Protected Shared Function createList(Of T As baseModel)(cp As CPBaseClass, sqlCriteria As String, sqlOrderBy As String) As List(Of T)
            Dim result As New List(Of T)
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim ignoreCacheNames As New List(Of String)
                Dim instanceType As Type = GetType(T)
                Dim contentName As String = derivedContentName(instanceType)
                If (cs.Open(contentName, sqlCriteria, sqlOrderBy)) Then
                    Dim instance As T
                    Do
                        instance = loadRecord(Of T)(cp, cs)
                        If (instance IsNot Nothing) Then
                            result.Add(instance)
                        End If
                        cs.GoNext()
                    Loop While cs.OK()
                End If
                cs.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' get the name of the record by it's id
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="recordId"></param>record
        ''' <returns></returns>
        Protected Shared Function getRecordName(Of T As baseModel)(cp As CPBaseClass, recordId As Integer) As String
            Try
                If (recordId > 0) Then
                    Dim instanceType As Type = GetType(T)
                    Dim tableName As String = derivedContentTableName(instanceType)
                    Dim cs As CPCSBaseClass = cp.CSNew()
                    If (cs.OpenSQL("select name from " & tableName & " where id=" & recordId.ToString())) Then
                        Return cs.GetText("name")
                    End If
                    cs.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return ""
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' get the name of the record by it's guid 
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ccGuid"></param>record
        ''' <returns></returns>
        Protected Shared Function getRecordName(Of T As baseModel)(cp As CPBaseClass, ccGuid As String) As String
            Try
                If (Not String.IsNullOrEmpty(ccGuid)) Then
                    Dim instanceType As Type = GetType(T)
                    Dim tableName As String = derivedContentTableName(instanceType)
                    Dim cs As CPCSBaseClass = cp.CSNew()
                    If (cs.OpenSQL("select name from " & tableName & " where ccguid=" & cp.Db.EncodeSQLText(ccGuid))) Then
                        Return cs.GetText("name")
                    End If
                    cs.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return ""
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' get the id of the record by it's guid 
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ccGuid"></param>record
        ''' <returns></returns>
        Protected Shared Function getRecordId(Of T As baseModel)(cp As CPBaseClass, ccGuid As String) As Integer
            Try
                If (Not String.IsNullOrEmpty(ccGuid)) Then
                    Dim instanceType As Type = GetType(T)
                    Dim tableName As String = derivedContentTableName(instanceType)
                    Dim cs As CPCSBaseClass = cp.CSNew()
                    If (cs.OpenSQL("select id from " & tableName & " where ccguid=" & cp.Db.EncodeSQLText(ccGuid))) Then
                        Return cs.GetInteger("id")
                    End If
                    cs.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return 0
        End Function
        '
        '====================================================================================================
        Protected Shared Function getCount(Of T As baseModel)(cp As CPBaseClass, sqlCriteria As String) As Integer
            Dim result As Integer = 0
            Try
                Dim instanceType As Type = GetType(T)
                Dim tableName As String = derivedContentTableName(instanceType)
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim sql As String = "select count(id) as cnt from " & tableName
                If (Not String.IsNullOrEmpty(sqlCriteria)) Then
                    sql &= " where " & sqlCriteria
                End If
                If (cs.OpenSQL(sql)) Then
                    result = cs.GetInteger("cnt")
                End If
                cs.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '====================================================================================================
        ''' <summary>
        ''' Temporary method to create a path for an uploaded. First, try the texrt value in the field. If it is empty, use this method to create the path,
        ''' append the filename to the end and save it to the field, and save the file there. This path starts with the tablename and ends with a slash.
        ''' </summary>
        ''' <param name="fieldName"></param>
        ''' <returns></returns>
        Protected Function getUploadPath(Of T As baseModel)(fieldName As String) As String
            Dim instanceType As Type = GetType(T)
            Dim tableName As String = derivedContentTableName(instanceType)
            Return tableName.ToLower() & "/" & fieldName.ToLower() & "/" & id.ToString().PadLeft(12, CChar("0")) & "/"
        End Function
    End Class
End Namespace
