using System;

// ReSharper disable IdentifierTypo

namespace SharpShell.Helpers
{
    /// <summary>
    ///     Helper for Win32, providing access to some of the key macros which are often used.
    /// </summary>
    public static class Win32Helper
    {
        /// <summary>
        ///     Gets the LoWord of an IntPtr.
        /// </summary>
        /// <param name="ptr">The int pointer.</param>
        /// <returns>The LoWord of an IntPtr.</returns>
        public static int LoWord(IntPtr ptr)
        {
            int @int = IntPtr.Size == 8 ? unchecked((int)ptr.ToInt64()) : ptr.ToInt32();
            return unchecked((short)@int);
        }

        /// <summary>
        ///     Gets the HiWord of an IntPtr.
        /// </summary>
        /// <param name="ptr">The int pointer.</param>
        /// <returns>The HiWord of an IntPtr.</returns>
        public static int HiWord(IntPtr ptr)
        {
            int @int = IntPtr.Size == 8 ? unchecked((int)ptr.ToInt64()) : ptr.ToInt32();
            return unchecked((short)((long)@int >> 16));
        }

        /// <summary>
        ///     Determines whether a value is an integer identifier for a resource.
        /// </summary>
        /// <param name="resource">The pointer to be tested whether it contains an integer resource identifier.</param>
        /// <returns>
        ///     <c>true</c> if all bits except the least 16 bits are zero, indicating 'resource' is an integer identifier for a
        ///     resource. Otherwise it is typically a pointer to a string.
        /// </returns>
        public static bool IS_INTRESOURCE(IntPtr resource)
        {
            return (uint)resource <= ushort.MaxValue;
        }
    }
}