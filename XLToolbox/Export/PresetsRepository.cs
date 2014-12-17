using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.IsolatedStorage;
using System.Xml.Serialization;
using System.Collections.ObjectModel;

namespace XLToolbox.Export
{
    /// <summary>
    /// Repository for export settings, is concerned with storing and
    /// retrieving a collection of <see cref="Presets"/>.
    /// </summary>
    [Serializable]
    public class PresetsRepository : IDisposable
    {
        #region Public properties

        public ObservableCollection<Preset> Presets { get; set; }

        #endregion

        #region Constructor

        public PresetsRepository()
            : base ()
        {
            // Must initialize the ExportSettings property, lest a null pointer
            // exception is thrown in the LoadSettings() method.
            Presets = new ObservableCollection<Export.Preset>();
            LoadPresets();
        }

        /// <summary>
        /// Creates a new Presets repository, loads previously saved presets
        /// and adds the <paramref name="addPreset"/> to the repository.
        /// </summary>
        /// <param name="addPreset">Preset to add to the repository.</param>
        public PresetsRepository(Preset addPreset)
            : this()
        {
            Presets.Add(addPreset);
        }

        #endregion

        #region Add and remove

        public void Add(Preset exportSettings)
        {
            Presets.Add(exportSettings);
        }

        public void Remove(Preset exportSettings)
        {
            Presets.Remove(exportSettings);
        }

        #endregion

        #region Load and save

        protected virtual void LoadPresets()
        {
            using (IsolatedStorageFile store = GetIsolatedStorageFile())
            {
                try
                {
                    if (store.FileExists(ISOSTOREFILENAME))
                    {
                        IsolatedStorageFileStream stream = store.OpenFile(ISOSTOREFILENAME,
                            System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
                        // The line below would fail if ExportSettings is null; however,
                        // the property is initialized in the constructor.
                        XmlSerializer serializer = new XmlSerializer(Presets.GetType());
                        Presets = serializer.Deserialize(stream) as ObservableCollection<Preset>;
                        stream.Close();
                    }
                    else
                    {
                        Presets = new ObservableCollection<Preset>();
                    }
                }
                catch 
                {
                    Presets = new ObservableCollection<Preset>();
                    // throw new StoreException("Cannot read export settings.", e);
                }
            }
        }

        protected virtual void SavePresets()
        {
            try
            {
                using (IsolatedStorageFile store = GetIsolatedStorageFile())
                {
                    IsolatedStorageFileStream stream = store.CreateFile(ISOSTOREFILENAME);
                    XmlSerializer serializer = new XmlSerializer(Presets.GetType());
                    serializer.Serialize(stream, Presets);
                    stream.Close();
                }
            }
            catch (Exception e)
            {
                throw new StoreException("Cannot write export settings.", e);
            }
        }

        private IsolatedStorageFile GetIsolatedStorageFile()
        {
            return IsolatedStorageFile.GetStore(
                IsolatedStorageScope.Roaming | IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null);
        }

        #endregion

        #region Disposal

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                SavePresets();
                _disposed = true;
            }
        }

        ~PresetsRepository()
        {
            Dispose(false);
        }

        #endregion

        #region Private fields

        bool _disposed;

        #endregion

        #region Private constants

        string ISOSTOREFILENAME = "export_settings.xml";

        #endregion
    }
}
