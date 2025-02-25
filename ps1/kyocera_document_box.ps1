# Step 1: Set printer name
#$printer_name = Read-Host -Prompt "Enter the name of the virtual printer (with FILE port)"
#$printer_ip = Read-Host -Prompt "Enter the IP of the installed device"
#$directory_path = Read-Host -Prompt "Enter the directory path where the document box files are stored"
$printer_name = "Kyocera TASKalfa 7054ci KX"
$printer_ip = "192.168.1.15"
$directory_path = "F:/boxes"
#Write-Host "FILE Printer: $printer_name"
#Write-Host "Printer IP: $printer_ip"
#Write-Host "Directory Path: $directory_path"

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class SendKeys {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

    public const byte key_CONTROL = 0x11;
    public const byte key_P = 0x50;
    public const byte key_ENTER = 0x0D;
    public const byte key_TAB = 0x09;
    public const uint press_KEYDOWN = 0x0000;
    public const uint press_KEYUP = 0x0002;

    public static void SendCtrlP() {
        keybd_event(key_CONTROL, 0, press_KEYDOWN, 0);
        keybd_event(key_P, 0, press_KEYDOWN, 0);
        keybd_event(key_P, 0, press_KEYUP, 0);
        keybd_event(key_CONTROL, 0, press_KEYUP, 0);
    }

    public static void SendEnter() {
        keybd_event(key_ENTER, 0, press_KEYDOWN, 0);
        keybd_event(key_ENTER, 0, press_KEYUP, 0);
    }

    public static void SendTab() {
        keybd_event(key_TAB, 0, press_KEYDOWN, 0);
        keybd_event(key_TAB, 0, press_KEYUP, 0);
    }
}
'@

function Create_PRN_files
{
    param (
        [string]$directory_path
    )
    # Step 2: Print the document to the FILE port and save the .prn file
    $files = Get-ChildItem -Path $directory_path -Recurse -File -Filter "*.pdf"
    Add-Type -AssemblyName "System.Windows.Forms"
    foreach ($file in $files)
    {
        $full_filename = $file.FullName
        $folder_path = [System.IO.Path]::GetDirectoryName($full_filename)
        Start-Process -FilePath "$full_filename" -ArgumentList "/t `"$printer_name`""

        Start-Sleep -Seconds 1
        [SendKeys]::SendCtrlP()

        Start-Sleep -Seconds 0.5
        [SendKeys]::SendEnter()

        Start-Sleep -Seconds 3
        for ($i = 0; $i -le 5; $i++) {
            [SendKeys]::SendTab()
        }
        [SendKeys]::SendEnter()

        # Step 3: Allow the Save As dialog to appear
        Start-Sleep -Seconds 1

        # Step 4: Use SendKeys to type the file path
        [System.Windows.Forms.SendKeys]::SendWait("$folder_path")

        # Step 5: Press Enter to save file
        [SendKeys]::SendEnter()

        Start-Sleep -Seconds 1
        for ($i = 0; $i -le 8; $i++) {
            [SendKeys]::SendTab()
        }
        Start-Sleep -Seconds 1
        [SendKeys]::SendEnter()

        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("%{F4}")

        Start-Sleep -Seconds 3
    }
}

# Step 6: Inject custom PJL commands for Document Boxes
function Inject_PJL_commands
{
    param (
        [string]$directory_path
    )

    $files = Get-ChildItem -Path $directory_path -Recurse -File -Filter "*.prn"

    foreach ($file in $files)
    {

        $full_name = $file.FullName
        $path_parts = $full_name -split '\\'
        $folder_name = $path_parts[-2]
        $filename = [System.IO.Path]::GetFileNameWithoutExtension($full_name)

        $formatted_box_number = "{0:D4}" -f ($folder_name -as [int])

        $content = Get-Content -Path $full_name
        $new_content = @()
        foreach ($line in $content)
        {
            #Write-Host "$line"
            if ($line -like "@PJL SET ECONOMODE=*")
            {
                $new_content += $line
                $new_content += "@PJL SET HOLD=KUSERBOX"
                $new_content += "@PJL SET KUSERBOXID=`"$formatted_box_number`""
                $new_content += "@PJL SET KUSERBOXPASSWORD=`"`""
            }
            elseif ($line -like "@PJL SET JOBNAME=*")
            {
                $new_content += "@PJL SET JOBNAME=`"$filename`""
            }
            else
            {
                $new_content += $line
            }
        }
        Set-Content -Path $full_name -Value $new_content
        Write-Host "Done Injecting $full_name"
    }
}

function LPR_files
{
    # Step 7: LPR the .prn file to actual printer
    Start-Process "lpr" -ArgumentList "-S $print_ip -P 9100 $prn_file_path"
}

Create_PRN_files -directory_path "$directory_path"
Inject_PJL_commands -directory_path "$directory_path"





