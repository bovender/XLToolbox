using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Export
{
    public class SettingsViewModelCollection : ViewModelCollection<Settings, SettingsViewModel>
    {
        public SettingsViewModelCollection(SettingsRepository settingsRepository)
            : base(settingsRepository.ExportSettings)
        { }

        protected override SettingsViewModel CreateViewModel(Settings model)
        {
            return new SettingsViewModel(model);
        }
    }
}
