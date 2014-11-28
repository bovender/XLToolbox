using System;
using System.Windows;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using XLToolbox.ExceptionHandler;
using XLToolbox.Excel.Instance;
using XLToolbox.Excel.ViewModels;
using XLToolbox.About;
using XLToolbox.Versioning;
using XLToolbox.SheetManager;

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
                        Bovender.Versioning.UpdaterViewModel uvm = new Bovender.Versioning.UpdaterViewModel(
                            new Updater()
                            );
                        System.Windows.Threading.Dispatcher d = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                        uvm.CheckForUpdateMessage.Sent += (object sender, MessageArgs<ProcessMessageContent> args) =>
                        {
                            Window view = args.Content.InjectInto<UpdaterProcessView>();
                            args.Content.ViewModel.ViewDispatcher = view.Dispatcher;
                            view.Show();
                            //d.Invoke(
                            //    new Action(() =>
                            //        args.Content.InjectAndShowInThread<UpdaterProcessView>()
                            //        )
                            //    );
                        };
                        uvm.CheckForUpdateCommand.Execute(null);
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
                ExceptionViewModel vm = new ExceptionViewModel(e);
                vm.InjectInto<ExceptionView>().ShowDialog();
            }
        }
    }
}
