Set WshShell = WScript.CreateObject("WScript.Shell")
Set shortcut = WshShell.CreateShortcut("C:\LuniaThings\Lunia Development\Client\Lunia Client.lnk")
shortcut.TargetPath = "C:\LuniaThings\Lunia Development\Client\LuniaClient.exe"
shortcut.Arguments = "teles1h"
shortcut.WorkingDirectory = "C:\LuniaThings\Lunia Development\Client"
shortcut.Save
WScript.Echo "Atalho criado com sucesso!"
