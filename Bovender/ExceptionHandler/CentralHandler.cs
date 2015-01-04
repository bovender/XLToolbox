/* CentralHandler.cs
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

namespace Bovender.ExceptionHandler
{
    /// <summary>
    /// This static class implements central exception management.
    /// </summary>
    /// <remarks>
    /// User entry points should try and catch exceptions which they
    /// may hand over to the <see cref="Manage"/> method. If an event
    /// handler is attached to the <see cref="ManageExceptionCallback"/>
    /// event, the event is raised. The event handler should display
    /// the exception to the user and set the event argument's
    /// <see cref="ManageExceptionEventArgs.IsHandled"/> property to
    /// true. If the Manage method does not find this property with a
    /// True value, it will throw an <see cref="UnhandeldException"/>
    /// exception itself which will contain the original exception as
    /// inner exception.
    /// </remarks>
    public static class CentralHandler
    {
        #region Events

        public static event EventHandler<ManageExceptionEventArgs> ManageExceptionCallback;

        #endregion

        #region Static methods

        /// <summary>
        /// Central exception managing method; can be called in try...catch
        /// statements of user entry points.
        /// </summary>
        /// <param name="origin">Object where the exception occurred.</param>
        /// <param name="e">Exception that occurred in <see cref="origin"/>.</param>
        /// <returns>True if central exception management was performed, false if not.
        /// If the exception was not managed, the calling method may want to rethrow
        /// the exception.</returns>
        public static bool Manage(object origin, Exception e)
        {
            ManageExceptionEventArgs args = new ManageExceptionEventArgs(e);
            if (ManageExceptionCallback != null)
            {
                ManageExceptionCallback(origin, args);
            }
            return args.IsHandled;
        }

        #endregion
    }
}
