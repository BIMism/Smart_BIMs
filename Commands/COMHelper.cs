using System;
using System.Runtime.InteropServices;

namespace Smart_BIMs.Commands
{
    public static class COMHelper
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public static object GetActiveObject(string progId)
        {
            Type t = Type.GetTypeFromProgID(progId);
            if (t == null) throw new Exception("COM Application not installed.");
            Guid clsid = t.GUID;
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            return obj;
        }
    }
}
