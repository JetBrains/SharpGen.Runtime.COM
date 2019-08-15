using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SharpGen.Runtime
{
    public static class ComObjectEx
    {
        /// <summary>
        /// Compares 2 COM objects and return true if the native pointer is the same.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns><c>true</c> if the native pointer is the same, <c>false</c> otherwise</returns>
        public static bool EqualsComObject<T>(T left, T right) where T : ComObject
        {
            return (Equals(left, right));
        }

        /// <summary>
        /// Queries a managed object for a particular COM interface support (This method is a shortcut to <see cref="ComObject.QueryInterface"/>)
        /// </summary>
        ///<typeparam name="T">The type of the COM interface to query</typeparam>
        /// <param name="comObject">The managed COM object.</param>
        ///<returns>An instance of the queried interface</returns>
        /// <msdn-id>ms682521</msdn-id>
        /// <unmanaged>IUnknown::QueryInterface</unmanaged>	
        /// <unmanaged-short>IUnknown::QueryInterface</unmanaged-short>
        public static T As<T>(object comObject) where T : ComObject
        {
            using (var tempObject = new ComObject(Marshal.GetIUnknownForObject(comObject)))
            {
                return tempObject.QueryInterface<T>();
            }
        }

        /// <summary>
        /// Queries a managed object for a particular COM interface support (This method is a shortcut to <see cref="ComObject.QueryInterface"/>)
        /// </summary>
        ///<typeparam name="T">The type of the COM interface to query</typeparam>
        /// <param name="iunknownPtr">The managed COM object.</param>
        ///<returns>An instance of the queried interface</returns>
        /// <msdn-id>ms682521</msdn-id>
        /// <unmanaged>IUnknown::QueryInterface</unmanaged>	
        /// <unmanaged-short>IUnknown::QueryInterface</unmanaged-short>
        public static T As<T>(IntPtr iunknownPtr) where T : ComObject
        {
            using (var tempObject = new ComObject(iunknownPtr))
            {
                return tempObject.QueryInterface<T>();
            }
        }

        /// <summary>
        /// Queries a managed object for a particular COM interface support.
        /// </summary>
        ///<typeparam name="T">The type of the COM interface to query</typeparam>
        /// <param name="comObject">The managed COM object.</param>
        ///<returns>An instance of the queried interface</returns>
        /// <msdn-id>ms682521</msdn-id>
        /// <unmanaged>IUnknown::QueryInterface</unmanaged>	
        /// <unmanaged-short>IUnknown::QueryInterface</unmanaged-short>
        public static T QueryInterface<T>(object comObject) where T : ComObject
        {
            using (var tempObject = new ComObject(Marshal.GetIUnknownForObject(comObject)))
            {
                return tempObject.QueryInterface<T>();
            }
        }

        /// <summary>
        /// Queries a managed object for a particular COM interface support.
        /// </summary>
        ///<typeparam name="T">The type of the COM interface to query</typeparam>
        /// <param name="comPointer">A pointer to a COM object.</param>
        ///<returns>An instance of the queried interface</returns>
        /// <msdn-id>ms682521</msdn-id>
        /// <unmanaged>IUnknown::QueryInterface</unmanaged>	
        /// <unmanaged-short>IUnknown::QueryInterface</unmanaged-short>
        public static T QueryInterfaceOrNull<T>(IntPtr comPointer) where T : ComObject
        {
            if (comPointer == IntPtr.Zero)
            {
                return null;
            }

            var guid = typeof(T).GetTypeInfo().GUID;
            IntPtr pointerT;
            var result = (Result)Marshal.QueryInterface(comPointer, ref guid, out pointerT);
            return (result.Failure) ? null : ComMarshallingHelpers.FromPointer<T>(pointerT);
        }
    }
}