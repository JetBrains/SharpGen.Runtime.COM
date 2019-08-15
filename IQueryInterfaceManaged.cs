namespace SharpGen.Runtime
{
    /// <summary>
    /// A common interface which provides QueryInterface methods to use from managed code
    /// </summary>
    public interface IQueryInterfaceManaged
    {
        /// <summary>
        /// Version of QueryInterface method to use from managed side
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>Queried object</returns>
        T QueryInterface<T>() where T : class, IUnknown;
        
        /// <summary>
        /// Version of QueryInterface method to use from managed side
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>Queried object or null</returns>
        T QueryInterfaceOrNull<T>() where T : class, IUnknown;
    }
}