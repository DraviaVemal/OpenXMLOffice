// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global;

namespace OpenXMLOffice.Excel {
    /// <summary>
    /// Represents the properties of a column in a worksheet.
    /// </summary>
    public class SpreadsheetProperties {
        #region Public Fields

        /// <summary>
        /// Spreadsheet settings
        /// </summary>
        public SpreadsheetSettings settings = new();

        /// <summary>
        /// Spreadsheet theme settings
        /// </summary>
        public ThemePallet theme = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings of a spreadsheet.
    /// </summary>
    public class SpreadsheetSettings {
    }
}