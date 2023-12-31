/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class TextBox : TextBoxBase
    {
        #region Public Constructors

        public TextBox(TextBoxSetting TextBoxSetting) : base(TextBoxSetting)
        {
        }

        #endregion Public Constructors

        #region Internal Methods

        internal A.Run GetTextBoxRun()
        {
            return base.GetTextBoxRun();
        }

        internal P.Shape GetTextBoxShape()
        {
            return base.GetTextBoxShape();
        }

        #endregion Internal Methods
    }
}