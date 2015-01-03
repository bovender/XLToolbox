using System.Windows;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Injects a view with a view model that is referenced in a ViewModelMessageContent,
    /// and shows the view modally.
    /// </summary>
    public class ShowViewDialogAction : ShowViewAction
    {
        protected override void ShowView(Window view)
        {
            view.ShowDialog();
        }
    }
}
