using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    /// <summary>
    /// COM Object Manager for automatic cleanup and memory management
    /// Implements RAII pattern to prevent memory leaks in Word automation
    /// </summary>
    public class ComObjectManager : IDisposable
    {
        private readonly List<object> _comObjects = new List<object>();
        private bool _disposed = false;

        /// <summary>
        /// Registers a COM object for automatic cleanup when disposing
        /// </summary>
        /// <typeparam name="T">Type of COM object</typeparam>
        /// <param name="comObject">The COM object to register</param>
        /// <returns>The same COM object for fluent usage</returns>
        public T Register<T>(T comObject) where T : class
        {
            if (comObject != null && !_disposed)
            {
                _comObjects.Add(comObject);
            }
            return comObject;
        }

        /// <summary>
        /// Disposes all registered COM objects in reverse order
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Release COM objects in reverse order (LIFO)
                    foreach (var obj in _comObjects.AsEnumerable().Reverse())
                    {
                        try
                        {
                            if (obj != null)
                            {
                                Marshal.FinalReleaseComObject(obj);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log but don't throw - cleanup should continue
                            Console.WriteLine($"COM cleanup warning: {ex.Message}");
                        }
                    }
                    _comObjects.Clear();
                }
                _disposed = true;
            }
        }

        /// <summary>
        /// Finalizer as backup - should not be called if Dispose is used properly
        /// </summary>
        ~ComObjectManager()
        {
            Dispose(false);
        }
    }
}