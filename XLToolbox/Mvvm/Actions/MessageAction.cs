﻿using System.Windows;
using Bovender.Mvvm.Actions;
using XLToolbox.Mvvm.Views;

namespace XLToolbox.Mvvm.Actions
{
    public class MessageAction : MessageActionBase
    {
        /// <summary>
        /// Returns a view that can bind to expected message contents.
        /// </summary>
        /// <returns>Descendant of Window.</returns>
        protected override Window CreateView()
        {
            return new MessageContentView();
        }
    }
}
