using System.Runtime.InteropServices;

namespace DwgAutoPrinter.App.Services;

internal static class ComInterop
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    public static bool TryGetActiveObject(string progId, out object? instance)
    {
        instance = null;
        var hr = CLSIDFromProgID(progId, out var clsid);
        if (hr < 0)
        {
            return false;
        }

        hr = GetActiveObject(ref clsid, IntPtr.Zero, out instance);
        return hr >= 0 && instance is not null;
    }
}
