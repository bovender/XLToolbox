using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Export
{
    public class PresetViewModelCollection : ViewModelCollection<Preset, PresetViewModel>
    {
        public PresetViewModelCollection(PresetsRepository settingsRepository)
            : base(settingsRepository.ExportSettings)
        { }

        protected override PresetViewModel CreateViewModel(Preset model)
        {
            return new PresetViewModel(model);
        }
    }
}
