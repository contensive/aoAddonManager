
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.Addons.AddonManager.Models
    Public Class personModel
        Inherits baseModel
        Implements ICloneable
        '
        '====================================================================================================
        '-- const
        Public Const contentName As String = "people"
        Public Const contentTableName As String = "ccmembers"
        Private Shadows Const contentDataSource As String = "default"
        '
        '====================================================================================================
        ' -- instance properties
        Public Property Address As String
        Public Property Address2 As String
        Public Property Admin As Boolean
        Public Property AdminMenuModeID As Integer
        Public Property AllowBulkEmail As Boolean
        Public Property AllowToolsPanel As Boolean
        Public Property AutoLogin As Boolean
        Public Property BillAddress As String
        Public Property BillAddress2 As String
        Public Property BillCity As String
        Public Property BillCompany As String
        Public Property BillCountry As String
        Public Property BillEmail As String
        Public Property BillFax As String
        Public Property BillName As String
        Public Property BillPhone As String
        Public Property BillState As String
        Public Property BillZip As String
        Public Property BirthdayDay As Integer
        Public Property BirthdayMonth As Integer
        Public Property BirthdayYear As Integer
        Public Property City As String
        Public Property Company As String
        Public Property Country As String
        Public Property CreatedByVisit As Boolean
        Public Property DateExpires As Date
        Public Property Developer As Boolean
        Public Property Email As String
        Public Property ExcludeFromAnalytics As Boolean
        Public Property Fax As String
        Public Property FirstName As String
        Public Property ImageFilename As String
        Public Property LanguageID As Integer
        Public Property LastName As String
        Public Property LastVisit As Date
        Public Property nickName As String
        Public Property NotesFilename As String
        Public Property OrganizationID As Integer
        Public Property Password As String
        Public Property Phone As String
        Public Property ResumeFilename As String
        Public Property ShipAddress As String
        Public Property ShipAddress2 As String
        Public Property ShipCity As String
        Public Property ShipCompany As String
        Public Property ShipCountry As String
        Public Property ShipName As String
        Public Property ShipPhone As String
        Public Property ShipState As String
        Public Property ShipZip As String
        Public Property State As String
        Public Property StyleFilename As String
        Public Property ThumbnailFilename As String
        Public Property Title As String
        Public Property Username As String
        Public Property Visits As Integer
        Public Property Zip As String
        '
        '====================================================================================================
        Public Overloads Shared Function add(cp As CPBaseClass) As personModel
            Return add(Of personModel)(cp)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordId As Integer) As personModel
            Return create(Of personModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordGuid As String) As personModel
            Return create(Of personModel)(cp, recordGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function createByName(cp As CPBaseClass, recordName As String) As personModel
            Return createByName(Of personModel)(cp, recordName)
        End Function
        '
        '====================================================================================================
        Public Overloads Sub save(cp As CPBaseClass)
            MyBase.save(cp)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, recordId As Integer)
            delete(Of personModel)(cp, recordId)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, ccGuid As String)
            delete(Of personModel)(cp, ccGuid)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Function createList(cp As CPBaseClass, sqlCriteria As String, Optional sqlOrderBy As String = "id") As List(Of personModel)
            Return createList(Of personModel)(cp, sqlCriteria, sqlOrderBy)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, recordId As Integer) As String
            Return baseModel.getRecordName(Of personModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, ccGuid As String) As String
            Return baseModel.getRecordName(Of personModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordId(cp As CPBaseClass, ccGuid As String) As Integer
            Return baseModel.getRecordId(Of personModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getCount(cp As CPBaseClass, sqlCriteria As String) As Integer
            Return baseModel.getCount(Of personModel)(cp, sqlCriteria)
        End Function
        '
        '====================================================================================================
        Public Overloads Function getUploadPath(fieldName As String) As String
            Return MyBase.getUploadPath(Of personModel)(fieldName)
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone(cp As CPBaseClass) As personModel
            Dim result As personModel = DirectCast(Me.Clone(), personModel)
            result.id = cp.Content.AddRecord(contentName)
            result.ccguid = cp.Utils.CreateGuid()
            result.save(cp)
            Return result
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function

    End Class
End Namespace
