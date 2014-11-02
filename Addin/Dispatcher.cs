using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using XLToolbox.Excel.Instance;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Mvvm.Views;
using XLToolbox.Mvvm.ViewModels;

namespace XLToolbox
{
    #region Commands enumeration

    enum Command
    {
        About,
        CheckForUpdates,
        ThrowError,
        SheetList
    };

    #endregion

    /// <summary>
    /// Central dispatcher for all UI-initiated XL Toolbox commands.
    /// </summary>
    /// <remarks>
    /// This static class is necessary to be able to handle unforeseen
    /// unhandled exceptions, which is otherwise not easy to achieve in
    /// VSTO projects, where the usual way via AppDomain.CurrentDomain.
    /// UnhandledException does not work.
    /// </remarks>
    static class Dispatcher
    {
        /// <summary>
        /// Central command dispatcher. This public method also contains
        /// the central error handler for user-friendly error messages.
        /// </summary>
        /// <param name="cmd">XL Toolbox command to execute.</param>
        public static void Execute(Command cmd)
        {
            try
            {
                switch (cmd)
                {
                    case Command.About:
                        AboutViewModel avm = new AboutViewModel();
                        avm.InjectInto<AboutView>().ShowDialog();
                        break;
                    case Command.CheckForUpdates:
                        // TODO: Implement MVVM here
                        // (new WindowCheckForUpdate()).ShowDialog();
                        break;
                    case Command.SheetList:
                        WorkbookViewModel wvm = new WorkbookViewModel(ExcelInstance.Application.ActiveWorkbook);
                        Workarounds.ShowModelessInExcel<WorkbookView>(wvm);
                        break;
                    case Command.ThrowError:
                        throw new InsufficientMemoryException();
                }
            }
            catch (Exception e)
            {
                // TODO: Implement global exception handler here
                /*
                ExceptionViewModel r = new ExceptionViewModel(Globals.ThisAddIn.Application, e);
                r.User = Properties.Settings.Default.UsersName;
                r.Email = Properties.Settings.Default.UsersEmail;
                r.CcUser = Properties.Settings.Default.CcUser;
                WindowRuntimeError w = new WindowRuntimeError(r);
                w.ShowDialog();
                Properties.Settings.Default.UsersName = r.User;
                Properties.Settings.Default.UsersEmail = r.Email;
                Properties.Settings.Default.CcUser = r.CcUser;
                Properties.Settings.Default.Save();
                */
            }
        }
    }
}
