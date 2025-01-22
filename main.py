import subprocess
import time
import sys
import os
import glob
import ctypes
import shutil
from win32 import win32gui
from win32com import client

# import win32.lib.win32con as win32con

from PIL import ImageGrab


# Force users to run the tool as administrator
def is_admin():
    try:
        # Check if the current process has administrator privileges
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def run_as_admin():
    if not is_admin():
        print(
            r"""      
   _________________________________________________________
 /                                                           \
|  Please right-click on this application and run with        |
|  Administrator Privileges                                   |
 \                                                           /
   =========================================================
                                       \
                                        \
                                         ,_     _
                                         |\\_,-~/
                                         / _  _ |    ,--.
                                        (  @  @ )   / ,-'
                                         \  _T_/-._( (
                                         /         `. \
                                        |         _  \ |
                                         \ \ ,  /      |
                                          || |-_\__   /
                                         ((_/`(____,-'
        
        """
        )
        os.system("pause")
        sys.exit()


run_as_admin()


# Debug mode
def warning_before(func):
    def inner(*args, **kwargs):
        while True:
            answer = input(
                "<Attention> Do not interact with desktop from now on. Ready? (y/n) "
            )
            if answer.lower() == "y":
                break
            elif answer.lower() == "n":
                sys.exit()
            else:
                continue
        func(*args, **kwargs)

    return inner


print(
    r"""
 _______________________________________
   Citrix GPU Cert Log Collection Tool 
                  v1.0 
 =======================================
"""
)

log_dir = "C:\\Users\\test\\Desktop\\Citrix_LOGS\\"


def is_window_fully_loaded(hwnd, title):
    """Check if a window with the specified title fully loaded"""
    return win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) == title


# Get Device Manager and DXDiag screenshots
@warning_before
def get_DM_DX_screenshots():
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Store shell object for later use
    shell = client.Dispatch("Shell.Application")

    # Change system language to English
    subprocess.run(
        [
            "powershell",
            "-command",
            r"Set-WinUserLanguageList -LanguageList en-US -Force",
        ],
        check=True,
    )

    # Show Desktop only
    shell.MinimizeAll()

    # Open Device Manager
    subprocess.run(
        [
            "powershell",
            "-command",
            r'Start-Process -FilePath "$env:SystemRoot\system32\hdwwiz.cpl"; Start-Sleep -Seconds 5; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{D}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{D}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}")',
        ],
        check=True,
    )

    # Wait for Device Manager window to open
    hwnd = win32gui.FindWindow(None, "Device Manager")
    while not is_window_fully_loaded(hwnd, "Device Manager"):
        time.sleep(1)

    # Open DXDiag
    subprocess.run(
        [
            "powershell",
            "-command",
            r'Start-Process -FilePath "$env:SystemRoot\system32\dxdiag.exe" -ArgumentList "/whql:on"; Start-Sleep -Seconds 10; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds 500; [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}")',
        ],
        check=True,
    )

    # Wait for DXDiag window to open
    hwnd = win32gui.FindWindow(None, "DirectX Diagnostic Tool")
    while not is_window_fully_loaded(hwnd, "DirectX Diagnostic Tool"):
        time.sleep(1)
    time.sleep(5)  # Allow some time for windows to fully load

    # Take screenshots and save
    try:
        screenshot = ImageGrab.grab()
        screenshot.save(os.path.join(log_dir, "device.png"))
        print("\n\033[32mdevice.png saved to Desktop\\Citrix_LOGS\033[0m\n")

        # Close Device Manager and DXDiag
        subprocess.run(
            [
                "powershell",
                "-Command",
                "Get-Process -Name mmc | Stop-Process -Force; Get-Process -Name dxdiag | Stop-Process -Force",
            ],
            shell=True,
        )

        # Restore all minimized windows
        shell.UndoMinimizeAll()

    except Exception as e:
        print(f"\n\033[31mError: Failed to save device.png: {e}\033[0m\n")
        shell.UndoMinimizeAll()


# Get Device Manager screenshot only
@warning_before
def get_DM_screenshot():
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Store shell object for later use
    shell = client.Dispatch("Shell.Application")

    # Change system language to English
    subprocess.run(
        [
            "powershell",
            "-command",
            r"Set-WinUserLanguageList -LanguageList en-US -Force",
        ],
        check=True,
    )

    # Show Desktop only
    shell.MinimizeAll()

    # Open Device Manager
    subprocess.run(
        [
            "powershell",
            "-command",
            r'Start-Process -FilePath "$env:SystemRoot\system32\hdwwiz.cpl"; Start-Sleep -Seconds 5; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{D}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{D}"); Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}")',
        ],
        check=True,
    )

    # Wait for Device Manager window to open
    hwnd = win32gui.FindWindow(None, "Device Manager")
    while not is_window_fully_loaded(hwnd, "Device Manager"):
        time.sleep(1)

    time.sleep(2)  # Allow some time for windows to fully load

    # Take screenshots and save
    try:
        screenshot = ImageGrab.grab()
        screenshot.save(os.path.join(log_dir, "Windows_crashdump-2.png"))
        print(
            "\n\033[32mWindows_crashdump-2.png saved to Desktop\\Citrix_LOGS\033[0m\n"
        )
        # Close Device Manager
        subprocess.run(
            [
                "powershell",
                "-Command",
                "Get-Process -Name mmc | Stop-Process -Force",
            ],
            shell=True,
        )

        # Restore all minimized windows
        shell.UndoMinimizeAll()

    except Exception as e:
        print(f"\n\033[31mError: Failed to save Windows_crashdump-2.png: {e}\033[0m\n")
        shell.UndoMinimizeAll()


# Run WinSAT
def get_WinSAT_1():
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    os.system("winsat formal -xml C:\\wsformal-1.xml")
    try:
        shutil.move("C:\\wsformal-1.xml", log_dir)
        print("\n\033[32mwsformal-1.xml saved to Desktop\\Citrix_LOGS\033[0m\n")
    except FileNotFoundError:
        print("\n\033[31mError: wsformal-1.xml file not found\033[0m\n")


# Restart the system and run WinSAT
def get_WinSAT_2():
    while True:
        answer = input("Do you want to restart the system now? (y/n) ")
        if answer.lower() == "y":
            os.system("shutdown /r /t 1")
            break
        elif answer.lower() == "n":
            break
        else:
            continue
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    os.system("winsat formal -xml C:\\wsformal-2.xml")
    try:
        shutil.move("C:\\wsformal-2.xml", log_dir)
        print("\n\033[32mwsformal-2.xml saved to Desktop\\Citrix_LOGS\033[0m\n")
    except FileNotFoundError:
        print("\n\033[31mError: wsformal-2.xml file not found\033[0m\n")


# Remove non-NVIDIA GPU drivers
def remove_nonNV_devices():
    try:
        # List Non-NVIDIA Display Adapters
        list_cmd = """
        $displayAdapters = Get-PnpDevice -Class "Display"
        $nonNvidiaAdapters = $displayAdapters | Where-Object { $_.FriendlyName -notlike "*NVIDIA*" }
        
        if ($nonNvidiaAdapters.Count -eq 0) {
            Write-Output "`nNon-NVIDIA display adapters not found. Exiting..."
            exit 0
        }
        
        Write-Output "`nNon-NVIDIA display adapters are listed below:"
        $nonNvidiaAdapters | ForEach-Object {
            Write-Output "  $($_.FriendlyName)"
        }
        
        Write-Output "`n"
        foreach ($adapter in $nonNvidiaAdapters) {
            Write-Output "Device: $($adapter.FriendlyName) ($($adapter.InstanceId))"
            $deviceId = $($adapter.InstanceId)
            pnputil /remove-device `"$deviceId`"
        }
        """
        subprocess.run(["powershell", "-Command", list_cmd], check=True)
        print("\n\033[32mDone!\033[0m\n")
    except subprocess.CalledProcessError as e:
        print(f"\n\033[31mExecution error: {e}\033[0m\n")
        raise
    except Exception as e:
        print(f"\n\033[31mUnexpetecd error: {e}\033[0m\n")
        raise


# Set Crashdump registry and disable auto-restart
def set_crashdump():
    # Set Crashdump registry
    subprocess.run(
        [
            "reg",
            "add",
            r"HKLM\SYSTEM\CurrentControlSet\Control\CrashControl",
            "/v",
            "CrashDumpEnabled",
            "/t",
            "REG_DWORD",
            "/d",
            "1",
            "/f",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    # Disable auto-restart
    subprocess.run(
        ["wmic", "recoveros", "set", "AutoReboot", "=", "False"],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    print("\n\033[32mRegistry and reboot policy set!\033[0m\n")


# Check license status and move license file
def check_NV_license():
    license_status = subprocess.run(
        [
            "powershell",
            "-Command",
            "nvidia-smi -q | Select-String 'License Status' | ForEach-Object { $_ -replace '.*: ', '' }",
        ],
        capture_output=True,
        text=True,
    )
    try:
        if license_status.stdout.strip().split()[0] == "Unlicensed":
            answer = input(
                "vGPU software is not licensed. Do you want to activate the license now? (y/n) "
            )
            while True:
                if answer.lower() == "y":
                    license_file = glob.glob(".\\client_configuration_token*.tok")
                    license_path = "C:\\Program Files\\NVIDIA Corporation\\vGPU Licensing\\ClientConfigToken"
                    if license_file:
                        shutil.move(license_file[0], license_path)
                        print(
                            "\n\033[32mLicense file (.tok) moved successfully\033[0m\n"
                        )
                        while True:
                            restart = input(
                                "Do you want to restart the system now? (y/n) "
                            )
                            if restart.lower() == "y":
                                os.system("shutdown /r /t 1")
                                break
                            elif restart.lower() == "n":
                                print("\n")
                                break
                            else:
                                continue
                        break
                    else:
                        print(
                            "\n\033[31mError: License file (.tok) not found in the current directory. Please check\033[0m\n"
                        )
                        break
                elif answer.lower() == "n":
                    print("\n")
                    break
                else:
                    answer = input(
                        "vGPU software is not licensed. Do you want to activate the license now? (y/n) "
                    )
        elif license_status.stdout.strip().split()[0] == "Licensed":
            print("\n\033[32mvGPU software is licensed!\033[0m\n")
        else:
            print("\n\033[31mError: License status is unknown\033[0m\n")
    except:
        print(
            "\n\033[31mError: Failed to check license status by nvidia-smi tool\033[0m\n"
        )


while True:
    answer = input(
        "  (1) Collect Device Manager & Dxdiag screenshot\n  (2) Collect Device Manager screenshot\n  (3) Collect 1st WinSAT log\n  (4) Collect 2nd WinSAT log (reboot required)\n  (5) Verify and activate vGPU license\n  (6) Uninstall non-NVIDIA GPU devices\n  (7) Set crashdump registry and disable auto-restart\n  (8) Exit\n\nSelect an action: "
    )
    if answer == "1":
        get_DM_DX_screenshots()
        continue
    elif answer == "2":
        get_DM_screenshot()
        continue
    elif answer == "3":
        get_WinSAT_1()
        continue
    elif answer == "4":
        get_WinSAT_2()
        continue
    elif answer == "5":
        check_NV_license()
        continue
    elif answer == "6":
        remove_nonNV_devices()
        continue
    elif answer == "7":
        set_crashdump()
        continue
    elif answer == "8":
        sys.exit()
    else:
        continue
