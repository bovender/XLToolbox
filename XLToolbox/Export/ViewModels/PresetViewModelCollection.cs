using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    public class PresetViewModelCollection : ViewModelCollection<Preset, PresetViewModel>
    {
        public PresetViewModelCollection(PresetsRepository settingsRepository)
            : base(settingsRepository.Presets)
        { }

        protected override PresetViewModel CreateViewModel(Preset model)
        {
            return new PresetViewModel(model);
        }
    }
}
