using Microsoft.VisualBasic;

namespace Contensive.Addons.AddonManager51 {
    public static class constants {
        // 
        public const string cr = Constants.vbCrLf;
        public const string legacyMethodInstallFromLibraryC41 = "{43810D33-2D31-4636-8A49-28F2305F4223}";
        public const string legacyMethodInstallFromPhysicalInstallPathC41 = "{7C96ED03-1266-4FC4-B08A-F290A8CD9CF5}";
        public const string rnUploadCollectionFile = "uploadFile";
        public const string rnBlockDependencies = "blockDependencies";
        public const string RequestNameButton = "button";
        public const string ButtonCancel = " Cancel ";
        public const string ButtonOK = " OK ";
        public const string RequestNameCollectionID = "collectionid";
        // 
        public const string iisRecycleAddonGuid = "{164C4B6F-2A89-4A0D-AAA0-7B34BB2DD77D}";
        // 
        public const int FieldTypeInteger = 1;       // An long number
        public const int FieldTypeText = 2;          // A text field (up to 255 characters)
        public const int FieldTypeLongText = 3;      // A memo field (up to 8000 characters)
        public const int FieldTypeBoolean = 4;       // A yes/no field
        public const int FieldTypeDate = 5;          // A date field
        public const int FieldTypeFile = 6;          // A filename of a file in the files directory.
        public const int FieldTypeLookup = 7;        // A lookup is a FieldTypeInteger that indexes into another table
        public const int FieldTypeRedirect = 8;      // creates a link to another section
        public const int FieldTypeCurrency = 9;      // A Float that prints in dollars
        public const int FieldTypeTextFile = 10;     // Text saved in a file in the files area.
        public const int FieldTypeImage = 11;        // A filename of a file in the files directory.
        public const int FieldTypeFloat = 12;        // A float number
        public const int FieldTypeAutoIncrement = 13; // long that automatically increments with the new record
        public const int FieldTypeManyToMany = 14;    // no database field - sets up a relationship through a Rule table to another table
        public const int FieldTypeMemberSelect = 15; // This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
        public const int FieldTypeCSSFile = 16;      // A filename of a CSS compatible file
        public const int FieldTypeXMLFile = 17;      // the filename of an XML compatible file
        public const int FieldTypeJavascriptFile = 18; // the filename of a javascript compatible file
        public const int FieldTypeLink = 19;           // Links used in href tags -- can go to pages or resources
        public const int FieldTypeResourceLink = 20;   // Links used in resources, link <img or <object. Should not be pages
        public const int FieldTypeHTML = 21;           // LongText field that expects HTML content
        public const int FieldTypeHTMLFile = 22;       // TextFile field that expects HTML content
        public const int FieldTypeMax = 22;
        // 
        public const string FieldDescriptorInteger = "Integer";
        public const string FieldDescriptorText = "Text";
        public const string FieldDescriptorLongText = "LongText";
        public const string FieldDescriptorBoolean = "Boolean";
        public const string FieldDescriptorDate = "Date";
        public const string FieldDescriptorFile = "File";
        public const string FieldDescriptorLookup = "Lookup";
        public const string FieldDescriptorRedirect = "Redirect";
        public const string FieldDescriptorCurrency = "Currency";
        public const string FieldDescriptorImage = "Image";
        public const string FieldDescriptorFloat = "Float";
        public const string FieldDescriptorManyToMany = "ManyToMany";
        public const string FieldDescriptorTextFile = "TextFile";
        public const string FieldDescriptorCSSFile = "CSSFile";
        public const string FieldDescriptorXMLFile = "XMLFile";
        public const string FieldDescriptorJavascriptFile = "JavascriptFile";
        public const string FieldDescriptorLink = "Link";
        public const string FieldDescriptorResourceLink = "ResourceLink";
        public const string FieldDescriptorMemberSelect = "MemberSelect";
        public const string FieldDescriptorHTML = "HTML";
        public const string FieldDescriptorHTMLFile = "HTMLFile";
        // 
        public const string FieldDescriptorLcaseInteger = "integer";
        public const string FieldDescriptorLcaseText = "text";
        public const string FieldDescriptorLcaseLongText = "longtext";
        public const string FieldDescriptorLcaseBoolean = "boolean";
        public const string FieldDescriptorLcaseDate = "date";
        public const string FieldDescriptorLcaseFile = "file";
        public const string FieldDescriptorLcaseLookup = "lookup";
        public const string FieldDescriptorLcaseRedirect = "redirect";
        public const string FieldDescriptorLcaseCurrency = "currency";
        public const string FieldDescriptorLcaseImage = "image";
        public const string FieldDescriptorLcaseFloat = "float";
        public const string FieldDescriptorLcaseManyToMany = "manytomany";
        public const string FieldDescriptorLcaseTextFile = "textfile";
        public const string FieldDescriptorLcaseCSSFile = "cssfile";
        public const string FieldDescriptorLcaseXMLFile = "xmlfile";
        public const string FieldDescriptorLcaseJavascriptFile = "javascriptfile";
        public const string FieldDescriptorLcaseLink = "link";
        public const string FieldDescriptorLcaseResourceLink = "resourcelink";
        public const string FieldDescriptorLcaseMemberSelect = "memberselect";
        public const string FieldDescriptorLcaseHTML = "html";
        public const string FieldDescriptorLcaseHTMLFile = "htmlfile";
    }
    /// <summary>
    /// legtacy styles
    /// </summary>

    public static class AfwStyles {
        public const string afwWidth10 = "afwWidth10";
        public const string afwWidth20 = "afwWidth20";
        public const string afwWidth30 = "afwWidth30";
        public const string afwWidth40 = "afwWidth40";
        public const string afwWidth50 = "afwWidth50";
        public const string afwWidth60 = "afwWidth60";
        public const string afwWidth70 = "afwWidth70";
        public const string afwWidth80 = "afwWidth80";
        public const string afwWidth90 = "afwWidth90";
        public const string afwWidth100 = "afwWidth100";
        //
        public const string afwWidth10px = "afwWidth10px";
        public const string afwWidth20px = "afwWidth20px";
        public const string afwWidth30px = "afwWidth30px";
        public const string afwWidth40px = "afwWidth40px";
        public const string afwWidth50px = "afwWidth50px";
        public const string afwWidth60px = "afwWidth60px";
        public const string afwWidth70px = "afwWidth70px";
        public const string afwWidth80px = "afwWidth80px";
        public const string afwWidth90px = "afwWidth90px";
        //
        public const string afwWidth100px = "afwWidth100px";
        public const string afwWidth200px = "afwWidth200px";
        public const string afwWidth300px = "afwWidth300px";
        public const string afwWidth400px = "afwWidth400px";
        public const string afwWidth500px = "afwWidth500px";
        //
        public const string afwMarginLeft100px = "afwMarginLeft100px";
        public const string afwMarginLeft200px = "afwMarginLeft200px";
        public const string afwMarginLeft300px = "afwMarginLeft300px";
        public const string afwMarginLeft400px = "afwMarginLeft400px";
        public const string afwMarginLeft500px = "afwMarginLeft500px";
        //
        public const string afwTextAlignRight = "afwTextAlignRight";
        public const string afwTextAlignLeft = "afwTextAlignLeft";
        public const string afwTextAlignCenter = "afwTextAlignCenter";
    }
}