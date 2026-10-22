# ===================================================================================================
#                              FUNCTION: Ensure_MailslotSender_Type
# ---------------------------------------------------------------------------------------------------
# Description:
#   Loads the C# MailslotSender type (kernel32 CreateFile/WriteFile) used for direct Windows
#   mailslot messaging. Idempotent: a PSTypeName guard makes repeat calls no-ops, and it is a
#   function (not a bare top-level block) so background jobs can load it too.
#
# Notes:
#   - Called once at script top; any job that receives its body can call it as well.
#   - Exposes [MailslotSender]::SendMailslotCommand(mailslotName, command).
# ===================================================================================================

function Ensure_MailslotSender_Type
{
	if (-not ([System.Management.Automation.PSTypeName]'MailslotSender').Type)
	{
		Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class MailslotSender {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteFile(
        IntPtr hFile,
        byte[] lpBuffer,
        uint nNumberOfBytesToWrite,
        out uint lpNumberOfBytesWritten,
        IntPtr lpOverlapped
    );
        
    [DllImport("kernel32.dll")]
    public static extern bool CloseHandle(IntPtr hObject);

    public static bool SendMailslotCommand(string mailslotName, string command) {
        const uint GENERIC_WRITE = 0x40000000;
        const uint FILE_SHARE_READ = 0x00000001;
        const uint OPEN_EXISTING = 3;

        IntPtr hFile = CreateFile(mailslotName, GENERIC_WRITE, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (hFile == new IntPtr(-1)) {
            return false;
        }

        byte[] data = Encoding.ASCII.GetBytes(command);
        uint bytesWritten;
        bool success = WriteFile(hFile, data, (uint)data.Length, out bytesWritten, IntPtr.Zero);
        CloseHandle(hFile);
        return success;
    }
}
"@
	}
}
