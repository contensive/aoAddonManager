namespace Contensive.AddonManager {
    /// <summary>
    /// Centralized constants for GUIDs, button names, request names, and CSS class references
    /// </summary>
    public static class Constants {
        /// <summary>
        /// Root XML node name for the collection list response
        /// </summary>
        public const string collectionListRootNode = "collectionlist";
        /// <summary>
        /// GUID for the addon manager library list layout
        /// </summary>
        public const string guidAddonManagerLibraryListCell = "{9767F464-3728-4B7D-904B-3442D7FD03BE}";
        /// <summary>
        /// Display name for the addon manager library list layout
        /// </summary>
        public const string nameAddonManagerLibraryLisCell = "Add-on Manager Library List Body Layout";
        /// <summary>
        /// Filename of the addon manager library list layout HTML template
        /// </summary>
        public const string pathFilenameAddonManagerLibraryLisCell = "AddonManagerLibraryBody.html";
        /// <summary>
        /// Form field name for collection file upload
        /// </summary>
        public const string rnUploadCollectionFile = "uploadFile";
        /// <summary>
        /// Form field name for the block-dependencies checkbox
        /// </summary>
        public const string rnBlockDependencies = "blockDependencies";
        /// <summary>
        /// Form field name for button actions
        /// </summary>
        public const string requestNameButton = "button";
        /// <summary>
        /// Cancel button label
        /// </summary>
        public const string buttonCancel = " Cancel ";
        /// <summary>
        /// OK button label
        /// </summary>
        public const string buttonOK = " OK ";
        /// <summary>
        /// Remove button form field name
        /// </summary>
        public const string buttonRemove = "button-remove";
        /// <summary>
        /// Form field name for collection ID
        /// </summary>
        public const string requestNameCollectionID = "collectionid";
        /// <summary>
        /// Admin framework CSS class name constants used by layout builder list forms
        /// </summary>
        public static class AfwStyles {
            /// <summary>50px width</summary>
            public const string afwWidth50px = "afwWidth50px";
            /// <summary>Left text alignment</summary>
            public const string afwTextAlignLeft = "afwTextAlignLeft";
            /// <summary>Center text alignment</summary>
            public const string afwTextAlignCenter = "afwTextAlignCenter";
        }
    }
}
