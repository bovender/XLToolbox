using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Injects a view with a view model that is referenced in a ViewModelMessageContent,
    /// and shows the view modally.
    /// </summary>
    public class ShowViewAction : MessageActionBase
    {
        #region Public properties

        public string Assembly { get; set; }
        public string View { get; set; }

        #endregion

        #region Overrides

        protected override Window CreateView()
        {
            object obj = Activator.CreateInstance(Assembly, View).Unwrap();
            Window view = obj as Window;
            if (view != null)
            {
                ViewModelMessageContent content = Content as ViewModelMessageContent;
                content.ViewModel.InjectInto(view);
                return view;
            }
            else
            {
                throw new ArgumentException(String.Format(
                    "Class name '{0}' in assembly '{1}' is not derived from Window.",
                    Assembly, View));
            }
        }

        protected override void ShowView(Window view)
        {
            view.Show();
        }

        #endregion
    }
}
