using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace XLToolbox
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            ButtonCheckForUpdate.Label = Strings.CheckForUpdates;
        }

        private void ButtonCheckForUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Dispatcher.Execute(Command.CheckForUpdates);
        }

        private void ButtonTestError_Click(object sender, RibbonControlEventArgs e)
        {
            Dispatcher.Execute(Command.ThrowError);
        }
    }
}
