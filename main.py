import subprocess
import time
import sys
import os
import glob
import ctypes
import shutil
from win32 import win32gui
from win32com import client
from colorama import init, Fore, Style
from PIL import ImageGrab

# Initialize colorama
init(autoreset=True)


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
                f"{Fore.YELLOW}<Attention> Do not interact with desktop from now on. Ready? (y/n){Style.RESET_ALL} "
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
        print(
            f"\n{Fore.GREEN}device.png saved to Desktop\\Citrix_LOGS{Style.RESET_ALL}\n"
        )

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
        print(f"{Fore.RED}Error: Failed to save device.png: {e}{Style.RESET_ALL}\n")
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
            f"\n{Fore.GREEN}Windows_crashdump-2.png saved to Desktop\\Citrix_LOGS{Style.RESET_ALL}\n"
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
        print(
            f"\n{Fore.RED}Error: Failed to save Windows_crashdump-2.png: {e}{Style.RESET_ALL}\n"
        )
        shell.UndoMinimizeAll()


# Run WinSAT
def get_WinSAT_1():
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    os.system("winsat formal -xml C:\\wsformal-1.xml")
    try:
        shutil.move("C:\\wsformal-1.xml", log_dir)
        print(
            f"\n{Fore.GREEN}wsformal-1.xml saved to Desktop\\Citrix_LOGS{Style.RESET_ALL}\n"
        )
    except FileNotFoundError:
        print(f"\n{Fore.RED}Error: wsformal-1.xml file not found{Style.RESET_ALL}\n")


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
        print(
            f"\n{Fore.GREEN}wsformal-2.xml saved to Desktop\\Citrix_LOGS{Style.RESET_ALL}\n"
        )
    except FileNotFoundError:
        print(f"\n{Fore.RED}Error: wsformal-2.xml file not found{Style.RESET_ALL}\n")


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
        print(
            f"\n{Fore.GREEN}Non-NVIDIA display adapters uninstalled!{Style.RESET_ALL}\n"
        )
    except subprocess.CalledProcessError as e:
        print(f"\n{Fore.RED}Execution error: {e}{Style.RESET_ALL}\n")
        raise
    except Exception as e:
        print(f"\n{Fore.RED}Unexpetecd error: {e}{Style.RESET_ALL}\n")
        raise


# Set Crashdump registry and disable auto-restart
def set_crashdump():
    restart_needed = False

    # Check current CrashDumpEnabled registry value
    result = subprocess.run(
        [
            "reg",
            "query",
            r"HKLM\SYSTEM\CurrentControlSet\Control\CrashControl",
            "/v",
            "CrashDumpEnabled",
        ],
        capture_output=True,
        text=True,
        check=True,
    )

    # Determine if registry needs to be set
    if "0x1" not in result.stdout:
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
        restart_needed = True

    # Disable auto-restart in all cases
    subprocess.run(
        ["wmic", "recoveros", "set", "AutoReboot", "=", "False"],
        check=True,
        stdout=subprocess.DEVNULL,
    )

    # Prompt for restart only if registry was changed
    if restart_needed:
        while True:
            answer = input(
                "Default registry changed to 1. Restart the system to take effect now? (y/n) "
            )
            if answer.lower() == "y":
                os.system("shutdown /r /t 1")
                break
            elif answer.lower() == "n":
                print("\n")
                break
            else:
                continue
    else:
        print(
            f"\n{Fore.GREEN}Crashdump registry and reboot policy set!{Style.RESET_ALL}\n"
        )


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
                            f"\n{Fore.GREEN}License file (.tok) moved successfully{Style.RESET_ALL}\n"
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
                            f"\n{Fore.RED}Error: License file (.tok) not found in the current directory. Please check{Style.RESET_ALL}\n"
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
            print(f"\n{Fore.GREEN}vGPU software is licensed!{Style.RESET_ALL}\n")
        else:
            print(f"\n{Fore.RED}Error: License status is unknown{Style.RESET_ALL}\n")
    except:
        print(
            f"\n{Fore.RED}Error: Failed to check license status by nvidia-smi tool{Style.RESET_ALL}\n"
        )


# Trigger BSOD on SUT VM
def trigger_VM_BSOD():
    """
    Command to trigger SUT VM BSOD on SUT shell:

    uuid=$(xe vm-list name-label='Windows 10-1' params=uuid | awk -F ': ' '{print $2}' | head -1)
    domain=$(list_domains | grep $uuid | awk '{print $1}')
    /usr/sbin/xen-hvmcrash $domain
    """
    # 底下指令已確認可成功運行
    # ssh root@192.168.x.x "/usr/sbin/xen-hvmcrash `list_domains | grep $(xe vm-list name-label='Windows10-1' params=uuid | awk -F ': ' '{print $2}' | head -1) | awk '{print $1}')'`"
    SUT_IP = input("Enter the SUT host IP: ")
    ssh_cmd = [
        "ssh",
        f"root@{SUT_IP}",
        f"/usr/sbin/xen-hvmcrash `list_domains | grep $(xe vm-list name-label='Windows 10-1' params=uuid | awk -F ': ' '{{print $2}}' | head -1) | awk '{{print $1}}'`",
    ]
    result = subprocess.run(ssh_cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print(
            f"\n{Fore.GREEN}Crash command sent to SUT VM successfully!{Style.RESET_ALL}\n"
        )
    else:
        print(f"\n{Fore.RED}Command failed.{Style.RESET_ALL}")
        print(f"\n{Fore.RED}Error: {result.stderr}{Style.RESET_ALL}")


# Copy TC VM
def copy_TC_VM():
    """
    Command to copy VM on SUT shell:

    uuid=$(xe vm-list name-label='Windows 10-1' params=uuid | awk -F ': ' '{print $2}' | head -1)

    """
    # 底下指令已確認可成功運行
    # ssh root@192.168.x.x "xe vm-copy vm=$(xe vm-list name-label='Windows 10-1' params=uuid | awk -F ': ' '{{print $2}}' | head -1) new-name-label='test'"
    while True:
        answer = input(
            f"{Fore.YELLOW}<Attention> Before copying VM, you need to shut down the base VM first. Ready? (y/n){Style.RESET_ALL} "
        )
        if answer.lower() == "y":
            break
        elif answer.lower() == "n":
            sys.exit()
        else:
            continue
    print("\n")

    while True:
        try:
            VM_count = int(input("Enter the number of VMs you want to copy: "))
            if VM_count < 1:
                continue
            break
        except ValueError:
            continue

    SUT_IP = input("Enter the SUT host IP: ")
    for i in range(1, VM_count + 1):
        print(f"\nNow copying VM #{i}...\n")
        ssh_cmd = [
            "ssh",
            f"root@{SUT_IP}",
            f"xe vm-copy vm=$(xe vm-list name-label='Windows 10-1' params=uuid | awk -F ': ' '{{print $2}}' | head -1) new-name-label='Windows 10-1 (Copy #{i})'",
        ]
        result = subprocess.run(ssh_cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"\n{Fore.GREEN}Copy VM #{i} successfully!{Style.RESET_ALL}\n")
        else:
            print(f"\n{Fore.RED}Copy #VM {i} failed.{Style.RESET_ALL}")
            print(f"\n{Fore.RED}Error: {result.stderr}{Style.RESET_ALL}")


while True:
    answer = input(
        " SUT:\n  (1) Collect Device Manager & Dxdiag screenshot\n  (2) Collect Device Manager screenshot\n  (3) Collect 1st WinSAT log\n  (4) Collect 2nd WinSAT log (reboot required)\n  (5) Verify and activate vGPU license\n  (6) Uninstall non-NVIDIA GPU devices\n  (7) Set crashdump registry and disable auto-restart\n TC:\n  (8) Trigger BSOD on SUT VM\n  (9) Copy VM\n  (Q) Quit\n\nSelect an action: "
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
        trigger_VM_BSOD()
        continue
    elif answer == "9":
        copy_TC_VM()
        continue
    elif answer.lower() == "q":
        sys.exit()
    else:
        continue
