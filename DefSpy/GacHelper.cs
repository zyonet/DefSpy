using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DefSpy
{

    internal struct ASSEMBLY_INFO
    {
        public uint cbAssemblyInfo;
        public uint dwAssemblyFlags;
        public ulong uliAssemblySizeInKB;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszCurrentAssemblyPathBuf;
        public uint cchBuf;
    }


    internal static class GacHelper
    {
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("e707dcde-d1cd-11d2-bab9-00c04f8eceae")]
        private interface IAssemblyCache
        {
            int Dummy1();

            [PreserveSig]
            IntPtr QueryAssemblyInfo(int flags, [MarshalAs(UnmanagedType.LPWStr)] string assemblyName, 
                ref ASSEMBLY_INFO assemblyInfo);

            int Dummy2();

            int Dummy3();

            int Dummy4();
        }
        [DllImport("fusion.dll")]
        private static extern IntPtr CreateAssemblyCache(out IAssemblyCache ppAsmCache, uint dwReserved);

        public static Tuple<string, bool> GetAssemblyPath(string assemblyName)
        {
            ASSEMBLY_INFO assemblyInfo = default(ASSEMBLY_INFO);
            assemblyInfo.cchBuf = 512;

            ASSEMBLY_INFO assemblyInfo2 = assemblyInfo;
            assemblyInfo2.pszCurrentAssemblyPathBuf = new string('\0', (int)assemblyInfo.cchBuf);
            IAssemblyCache ppAsmCache;
            IntPtr value = CreateAssemblyCache(out ppAsmCache, 0);
            if (value == IntPtr.Zero)
            {
                value = ppAsmCache.QueryAssemblyInfo(1, assemblyName, ref assemblyInfo2);
                if (value != IntPtr.Zero)
                {
                    return new Tuple<string, bool>(assemblyInfo2.pszCurrentAssemblyPathBuf, false);
                }
                return new Tuple<string, bool>(assemblyInfo2.pszCurrentAssemblyPathBuf, true);
            }
            if (Debugger.IsAttached)
            {
                Marshal.ThrowExceptionForHR(value.ToInt32());
            }
            return new Tuple<string, bool>(assemblyInfo2.pszCurrentAssemblyPathBuf, false);
        }
    }
}
