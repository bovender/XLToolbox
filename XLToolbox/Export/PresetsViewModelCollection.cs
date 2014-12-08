using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Export
{
    public class PresetsViewModelCollection : ViewModelCollection<Preset, PresetsViewModel>
    {
        public PresetsViewModelCollection(PresetsRepository settingsRepository)
            : base(settingsRepository.ExportSettings)
        { }

        protected override PresetsViewModel CreateViewModel(Preset model)
        {
            return new PresetsViewModel(model);
        }
    }
}
