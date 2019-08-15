using System;

namespace SharpGen.Runtime
{
    /// <summary>
    /// Base class for COM callbacks which implements <see cref="IUnknown"/> and supports QI for regular .NET inheritance as well
    /// </summary>
    public class ComCallbackBase : CallbackBase, IUnknown
    {
        public T QueryInterface<T>() where T : class, IUnknown
        {
            if (this is T result) 
                return result;
            throw new SharpGenException(Result.NoInterface);
        }

        public T QueryInterfaceOrNull<T>() where T : class, IUnknown
        {
            return this as T;
        }

        public Result QueryInterface(Guid guid, out IntPtr output)
        {
            output = ((ICallbackable) this).Shadow.Find(guid);

            if (output == IntPtr.Zero)
            {
                return Result.NoInterface.Code;
            }

            AddRef();
            return Result.Ok.Code;
        }
    }
}