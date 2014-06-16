using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Core
{
    public class WorkbookStorage
    {
        private string _context;
        public Workbook workbook { get; private set; }

        /// <summary>
        /// Instantiates the class and associates it with the active workbook of the
        /// given application.
        /// </summary>
        /// <param name="application">Instance of an Excel application.</param>
        public WorkbookStorage(Application application)
        {
            workbook = application.ActiveWorkbook;
        }

        /// <summary>
        /// Instantiates the class and associates it with a workbook.
        /// </summary>
        /// <param name="workbook">Workbook object to associate the storage with.</param>
        public WorkbookStorage(Workbook workbook)
        {
            this.workbook = workbook;
        }


        /// <summary>
        /// Sets the context of the current storage object. This may be a
        /// worksheet name or an empty string for the global context of
        /// the workbook.
        /// </summary>
        // TODO: Deal with worksheets that are renamed or deleted while the WorkbookStorage is instantiated.
        public string Context {
            get
            {
                return _context;
            }
            set
            {
                /// If context is not an empty string, it denotes a worksheet
                /// in the associated workbook. The setter will test if the workbook
                /// does indeed contain such a worksheet, throwing an exception if not.
                if (value.Length > 0) {
                    try {
                        object o = workbook.Sheets[value];
                        _context = value;
                    }
                    catch (System.Runtime.InteropServices.COMException e) {
                        throw new WorkbookStorageException("Invalid storage context", e);
                    }
                }
                else
                {
                    _context = value;
                }
            }
        }

        /// <summary>
        /// Retrieves an integer from the storage, given a key.
        /// </summary>
        /// <param name="key">Key to look up.</param>
        /// <returns>Integer value</returns>
        public int Retrieve(string key)
        {
            throw new NotImplementedException();
        }

        public void Store(string key, int i)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Sets the active sheet of the current workbook as the context.
        /// </summary>
        public void UseActiveSheet()
        {
            Context = workbook.ActiveSheet.Name;
        }
    }
}
