<# : batch portion (begins PowerShell multiline comment block)
@echo off & setlocal

echo ------------------------------------------------
echo HID-E
echo Made by Arimodu
echo License: MIT
echo ------------------------------------------------

:start

Echo Waiting for ctrl-T...

rem # re-launch self with PowerShell interpreter
powershell -noprofile "iex (${%~f0} | out-string)"

echo Done.
Echo Looping to start...

goto start
: end batch / begin PowerShell chimera #>



$cSource = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class Clicker
{
//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct INPUT
{ 
    public int        type; // 0 = INPUT_MOUSE,
                            // 1 = INPUT_KEYBOARD
                            // 2 = INPUT_HARDWARE
    public MOUSEINPUT mi;
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT
{
    public int    dx ;
    public int    dy ;
    public int    mouseData ;
    public int    dwFlags;
    public int    time;
    public IntPtr dwExtraInfo;
}

//This covers most use cases although complex mice may have additional buttons
//There are additional constants you can use for those cases, see the msdn page
const int MOUSEEVENTF_MOVED      = 0x0001 ;
const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
const int MOUSEEVENTF_LEFTUP     = 0x0004 ;
const int MOUSEEVENTF_RIGHTDOWN  = 0x0008 ;
const int MOUSEEVENTF_RIGHTUP    = 0x0010 ;
const int MOUSEEVENTF_MIDDLEDOWN = 0x0020 ;
const int MOUSEEVENTF_MIDDLEUP   = 0x0040 ;
const int MOUSEEVENTF_WHEEL      = 0x0080 ;
const int MOUSEEVENTF_XDOWN      = 0x0100 ;
const int MOUSEEVENTF_XUP        = 0x0200 ;
const int MOUSEEVENTF_ABSOLUTE   = 0x8000 ;

const int screen_length = 0x10000 ;

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
[System.Runtime.InteropServices.DllImport("user32.dll")]
extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

public static void LeftClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}

public static void RightClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_RIGHTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}

}
'@
Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing


# import GetAsyncKeyState()
Add-Type user32_dll @'
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
'@ -namespace System

# for Keys object constants
Add-Type -As System.Windows.Forms

function keyPressed($key) {
    return [user32_dll]::GetAsyncKeyState([Windows.Forms.Keys]::$key) -band 32768
}

while ($true) {
    $ctrl = keyPressed "ControlKey"
    $key = keyPressed "T"
    if ($ctrl -and $key) { break }
    start-sleep -milliseconds 40
}

$wshell = New-Object -ComObject wscript.shell;
Sleep 1


# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa202943(v=office.10)?redirectedfrom=MSDN

$wshell.SendKeys('^p')
Write-host "Sending CTRL + P"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('~')
Write-host "Sending ENTER"
Sleep 0.5

$wshell.SendKeys('^p')
Write-host "Sending CTRL + P"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('{DOWN}')
Write-host "Sending DOWN"
$wshell.SendKeys('~')
Write-host "Sending ENTER"
Sleep 1
$wshell.SendKeys('{ESC}')
Write-host "Sending ESC"


#Send a click at a specified point
$Pos = [System.Windows.Forms.Cursor]::Position
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point( ($Pos.X) , ($Pos.Y) )
[Clicker]::LeftClickAtPoint(($Pos.X) , ($Pos.Y))
Write-host "Sending Left Click at cursor position"
