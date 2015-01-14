/* EnumProvider.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Bovender.Mvvm.ViewModels;

namespace Bovender.Mvvm
{
    /// <summary>
    /// Facilitates WPF data binding to enums by providing an enumeration of
    /// Choices, read/write access to a string representation (which may be
    /// localized in derived classes), and the type-safe enum value itself.
    /// </summary>
    /// <remarks>
    /// <para>
    /// To bind to a ComboBox to an EnumProvider property, use:
    ///     <code>
    ///         <ComboBox 
    ///             ItemsSource="{Binding MyEnumProviderProperty.Choices}"
    ///             ToolTip="{Binding MyEnumProviderProperty.ToolTip}"
    ///             SelectedItem="{Binding MyEnumProviderProperty.SelectedItem}"
    ///         />
    ///     </code>
    /// </para>
    /// <para>
    /// To make use of per-item enabled states and tool tips, it is helpful
    /// to define a generic style in a central resource dictionary:
    ///     <code>
    ///         <Style TargetType="{x:Type ComboBoxItem}">
    ///             <Setter Property="Control.ToolTip" Value="{Binding Path=ToolTip, Mode=OneWay}" />
    ///             <Setter Property="IsEnabled" Value="{Binding Path=IsEnabled, Mode=OneWay}" />
    ///         </Style>
    ///     </code>
    /// </para>
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
                return _selectedItem.Value;
            }
            set
            {
                _selectedItem = GetViewModel(value);
                AllPropertiesChanged();
            }
        }

        public EnumViewModel<T> SelectedItem
        {
            get
            {
                return _selectedItem;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(
                        "SelecteItem cannot be null because enums are not nullable.");
                }
                _selectedItem = value;
                AllPropertiesChanged();
            }
        }

        public string ToolTip
        {
            get
            {
                return SelectedItem.ToolTip;
            }
        }

        /// <summary>
        /// Returns an array of enum view models that represent the enum members.
        /// </summary>
        public IEnumerable<EnumViewModel<T>> Choices
        {
            get
            {
                if (_choices == null)
                {
                    _choices = new Collection<EnumViewModel<T>>();
                    foreach (T member in Enum.GetValues(typeof(T)))
                    {
                        EnumViewModel<T> vm = new EnumViewModel<T>(
                            member,
                            GetDescription(member),
                            GetTooltip(member));
                        vm.PropertyChanged += EnumViewModel_PropertyChanged;
                        _choices.Add(vm);
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
            return null;
        }

        void EnumViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender == SelectedItem)
            {
                OnPropertyChanged("SelectedItem");
            }
            OnPropertyChanged("Choices");
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Returns the view model for a given member. The view model
        /// may be used to enable/disable the member, or set a different
        /// description or tooltip.
        /// </summary>
        /// <param name="member">Member whose view model to return.</param>
        /// <returns>Instance of EnumViewModel </returns>
        public EnumViewModel<T> GetViewModel(T member)
        {
            return Choices.FirstOrDefault(item => item.Value.Equals(member));
        }

        #endregion

        #region Private methods

        private void AllPropertiesChanged()
        {
                OnPropertyChanged("SelectedItem");
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
        private EnumViewModel<T> _selectedItem;
        private Collection<EnumViewModel<T>> _choices;

        #endregion
    }
}
