using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Data;

namespace Bovender.Mvvm
{
    /// <summary>
    /// Facilitates WPF data binding to enums by providing an enumeration of
    /// Choices, read/write access to a string representation (which may be
    /// localized in derived classes), and the type-safe enum value itself.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Since generic type parameters cannot be enums, the workaround
    /// "struct, IConvertible" is used here as suggested in
    /// http://stackoverflow.com/q/79126/270712
    /// </para>
    /// </remarks>
    [Serializable]
    public class EnumProvider<T> : INotifyPropertyChanged where T: struct, IConvertible
    {
        #region Public properties

        public T AsEnum
        {
            get
            {
                return _enum;
            }
            set
            {
                _enum = value;
                AllPropertiesChanged();
            }
        }

        public string AsString
        {
            get
            {
                return GetDescription(AsEnum);
            }
            set
            {
                _enum = StringToEnum(value);
                AllPropertiesChanged();
            }
        }

        public string Tooltip
        {
            get
            {
                return GetTooltip(AsEnum);
            }
        }

        /// <summary>
        /// Returns an array of strings that represent the enum members.
        /// </summary>
        public IEnumerable<string> Choices
        {
            get
            {
                if (_choices == null)
                {
                    _choices = new Collection<string>();
                    foreach (T member in Enum.GetValues(typeof(T)))
                    {
                        _choices.Add(GetDescription(member));
                    }
                }
                return _choices;
            }
        }

        #endregion

        #region Constructors

        public EnumProvider() {}

        public EnumProvider(T initialValue)
            : this()
        {
            _enum = initialValue;
        }

        #endregion

        #region Virtual methods

        /// <summary>
        /// Returns a display string for a given enum value. In the base class,
        /// this is the description attribute, if present. Derived classes may
        /// override this to return localized strings.
        /// </summary>
        /// <param name="forValue">Enum value for which to return
        /// a display string.</param>
        /// <returns>Display string</returns>
        /// <remarks>See http://stackoverflow.com/a/1799401/270712
        /// for description of attribute accession.</remarks>
        protected virtual string GetDescription(T member)
        {
            Type type = typeof(T);
            MemberInfo[] memberInfo = type.GetMember(member.ToString());
            object[] attributes = memberInfo[0].GetCustomAttributes(
                typeof(DescriptionAttribute), false);
            if (attributes.Length > 0 && attributes[0] is DescriptionAttribute)
            {
                return ((DescriptionAttribute)attributes[0]).Description;
            }
            else
            {
                return member.ToString();
            }
        }

        /// <summary>
        /// Returns a tooltip for the given enum member. Derived
        /// classes may override this method to return localized
        /// tooltips.
        /// </summary>
        /// <param name="member">Enum member for which to return
        /// a tooltip.</param>
        /// <returns>Tooltip string (may be localized in derived
        /// classes).</returns>
        protected virtual string GetTooltip(T member)
        {
            return GetDescription(member);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Finds the index of the description string <paramref name="text"/>
        /// in the Choices array.
        /// </summary>
        /// <param name="text">Description string to search for.</param>
        /// <returns>Index of the description string in the Choices array.</returns>
        private int GetIndex(string text)
        {
            if (Choices.Contains(text))
            {
                return _choices.IndexOf(text);
            }
            else
            {
                throw new ArgumentException(
                    "Enumeration descriptions do not contain " + text);
            }
        }

        private T StringToEnum(string text)
        {
            int index = GetIndex(text);
            T[] values = ((T[])System.Enum.GetValues(typeof(T)));
            return values[index];
        }

        private void AllPropertiesChanged()
        {
                OnPropertyChanged("AsString");
                OnPropertyChanged("AsEnum");
                OnPropertyChanged("Tooltip");
        }

        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Private fields

        private T _enum;
        private Collection<string> _choices;

        #endregion
    }
}
