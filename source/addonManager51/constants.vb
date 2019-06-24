
Option Explicit On
Option Strict On

Namespace Contensive.Addons.AddonManager51
    Public Module constants
        '
        Public Const cr As String = vbCrLf
        Public Const legacyMethodInstallFromLibraryC41 As String = "{43810D33-2D31-4636-8A49-28F2305F4223}"
        Public Const legacyMethodInstallFromPhysicalInstallPathC41 As String = "{7C96ED03-1266-4FC4-B08A-F290A8CD9CF5}"
        Public Const rnUploadCollectionFile As String = "uploadFile"
        Public Const rnUploadReinstallDependencies As String = "uploadInstallDependencies"
        Public Const RequestNameButton As String = "button"
        Public Const ButtonCancel As String = " Cancel "
        Public Const ButtonOK As String = " OK "
        Public Const RequestNameCollectionID As String = "collectionid"
        '
        Public Const iisRecycleAddonGuid As String = "{164C4B6F-2A89-4A0D-AAA0-7B34BB2DD77D}"
        '
        Public Const FieldTypeInteger As Integer = 1       ' An long number
        Public Const FieldTypeText As Integer = 2          ' A text field (up to 255 characters)
        Public Const FieldTypeLongText As Integer = 3      ' A memo field (up to 8000 characters)
        Public Const FieldTypeBoolean As Integer = 4       ' A yes/no field
        Public Const FieldTypeDate As Integer = 5          ' A date field
        Public Const FieldTypeFile As Integer = 6          ' A filename of a file in the files directory.
        Public Const FieldTypeLookup As Integer = 7        ' A lookup is a FieldTypeInteger that indexes into another table
        Public Const FieldTypeRedirect As Integer = 8      ' creates a link to another section
        Public Const FieldTypeCurrency As Integer = 9      ' A Float that prints in dollars
        Public Const FieldTypeTextFile As Integer = 10     ' Text saved in a file in the files area.
        Public Const FieldTypeImage As Integer = 11        ' A filename of a file in the files directory.
        Public Const FieldTypeFloat As Integer = 12        ' A float number
        Public Const FieldTypeAutoIncrement As Integer = 13 'long that automatically increments with the new record
        Public Const FieldTypeManyToMany As Integer = 14    ' no database field - sets up a relationship through a Rule table to another table
        Public Const FieldTypeMemberSelect As Integer = 15 ' This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
        Public Const FieldTypeCSSFile As Integer = 16      ' A filename of a CSS compatible file
        Public Const FieldTypeXMLFile As Integer = 17      ' the filename of an XML compatible file
        Public Const FieldTypeJavascriptFile As Integer = 18 ' the filename of a javascript compatible file
        Public Const FieldTypeLink As Integer = 19           ' Links used in href tags -- can go to pages or resources
        Public Const FieldTypeResourceLink As Integer = 20   ' Links used in resources, link <img or <object. Should not be pages
        Public Const FieldTypeHTML As Integer = 21           ' LongText field that expects HTML content
        Public Const FieldTypeHTMLFile As Integer = 22       ' TextFile field that expects HTML content
        Public Const FieldTypeMax As Integer = 22
        '
        Public Const FieldDescriptorInteger = "Integer"
        Public Const FieldDescriptorText = "Text"
        Public Const FieldDescriptorLongText = "LongText"
        Public Const FieldDescriptorBoolean = "Boolean"
        Public Const FieldDescriptorDate = "Date"
        Public Const FieldDescriptorFile = "File"
        Public Const FieldDescriptorLookup = "Lookup"
        Public Const FieldDescriptorRedirect = "Redirect"
        Public Const FieldDescriptorCurrency = "Currency"
        Public Const FieldDescriptorImage = "Image"
        Public Const FieldDescriptorFloat = "Float"
        Public Const FieldDescriptorManyToMany = "ManyToMany"
        Public Const FieldDescriptorTextFile = "TextFile"
        Public Const FieldDescriptorCSSFile = "CSSFile"
        Public Const FieldDescriptorXMLFile = "XMLFile"
        Public Const FieldDescriptorJavascriptFile = "JavascriptFile"
        Public Const FieldDescriptorLink = "Link"
        Public Const FieldDescriptorResourceLink = "ResourceLink"
        Public Const FieldDescriptorMemberSelect = "MemberSelect"
        Public Const FieldDescriptorHTML = "HTML"
        Public Const FieldDescriptorHTMLFile = "HTMLFile"
        '
        Public Const FieldDescriptorLcaseInteger = "integer"
        Public Const FieldDescriptorLcaseText = "text"
        Public Const FieldDescriptorLcaseLongText = "longtext"
        Public Const FieldDescriptorLcaseBoolean = "boolean"
        Public Const FieldDescriptorLcaseDate = "date"
        Public Const FieldDescriptorLcaseFile = "file"
        Public Const FieldDescriptorLcaseLookup = "lookup"
        Public Const FieldDescriptorLcaseRedirect = "redirect"
        Public Const FieldDescriptorLcaseCurrency = "currency"
        Public Const FieldDescriptorLcaseImage = "image"
        Public Const FieldDescriptorLcaseFloat = "float"
        Public Const FieldDescriptorLcaseManyToMany = "manytomany"
        Public Const FieldDescriptorLcaseTextFile = "textfile"
        Public Const FieldDescriptorLcaseCSSFile = "cssfile"
        Public Const FieldDescriptorLcaseXMLFile = "xmlfile"
        Public Const FieldDescriptorLcaseJavascriptFile = "javascriptfile"
        Public Const FieldDescriptorLcaseLink = "link"
        Public Const FieldDescriptorLcaseResourceLink = "resourcelink"
        Public Const FieldDescriptorLcaseMemberSelect = "memberselect"
        Public Const FieldDescriptorLcaseHTML = "html"
        Public Const FieldDescriptorLcaseHTMLFile = "htmlfile"
    End Module
End Namespace
