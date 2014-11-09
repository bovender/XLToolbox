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

        public static void Manage(object origin, Exception e)
        {
            ManageExceptionEventArgs args = new ManageExceptionEventArgs(e);
            if (ManageExceptionCallback != null)
            {
                ManageExceptionCallback(origin, args);
            }
            if (!args.IsHandled)
            {
                throw new UnhandledException("No central handler managed the exception", e);
            }
        }

        #endregion
    }
}
