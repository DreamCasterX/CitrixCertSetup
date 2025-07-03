#!/usr/bin/env bash


# CREATOR: Mike Lu (klu7@lenovo.com)
# CHANGE DATE: 3/11/2025
__version__="1.1"


# Quick Setup For Citrix Cert testing**
# GPUs that support multiple display mode are: A40/L40/L40S/5000 Ada/6000 Ada/A5000/A5500/A6000

# [Note]
# Copy 'displaymodeselector' and 'NVIDIA vGPU driver (.rpm)' to the same path as this script

# ====================================================================

# Ensure the user is running the script as root
if [ "$EUID" -ne 0 ]; then 
    echo "⚠️ Please login as root to run this script"
    exit 1
fi 


# Check the existence of required files
[[ ! -f ./displaymodeselector ]] && echo -e "\n❌ ERROR: NVIDIA display mode selector tool (displaymodeselector) not found in the current directory" && exit 1
chmod 755 ./displaymodeselector
[[ ! `find ./ -name "NVIDIA*.rpm"` ]] && echo -e "\n❌ ERROR: NVIDIA driver package not found in the current directory" && exit 1
[[ $(ls ./*.rpm | wc -l) > 1 ]] && echo -e "\n❌ ERROR: More than one NVIDIA driver packages found in the current directory! \n" && exit 1

# Set GPU mode to Compute for supported GPUs
if [[ ! `lsmod | grep -i nvidia` ]]; then
    echo
    echo "╭──────────────────────────────────╮"
    echo "│    Setting GPU mode to Compute   │"
    echo "│                                  │"
    echo "╰──────────────────────────────────╯"
    echo
    gpu_mode=`./displaymodeselector --version | grep 'GPU Mode' | awk -F ': ' '{print $2}'` 

    if [[ $gpu_mode == "Compute" ]]; then
        echo -e "\n✅ GPU mode set to Compute"
    elif [[ $gpu_mode == "Physical display disabled" ]]; then  # for suppotred GPUs
        ./displaymodeselector --gpumode
        gpu_mode= `./displaymodeselector --version | grep 'GPU Mode' | awk -F ': ' '{print $2}'`
        if [[ $gpu_mode == "Compute" ]]; then
            echo -e "\n✅ GPU mode set to Compute"
        else
            echo -e "\n❌ ERROR: Failed to set GPU mode to Compute" && exit 1
        fi
    elif [[ $gpu_mode == "N/A" ]]; then  # for unsupported GPUs
        echo -e "\n✅ GPU does not support multiple display mode, skipping..."
    fi
fi	

# Install NV vGPU driver
echo
echo "╭───────────────────────────────────────╮"
echo "│     Installing NVIDIA vGPU driver     │"
echo "│                                       │"
echo "╰───────────────────────────────────────╯"
echo
if [[ ! `lsmod | grep -i nvidia` ]]; then
    if ! `rpm -qa | grep -i 'nvidia'`; then
        rpm -ivh ./NVIDIA*.rpm
        [[ $? == 0 ]] && systemctl reboot || { echo "❌ ERROR: Failed to install NV driver"; exit 1; }
    fi
else
    echo -e "\n✅ NV vGPU driver is installed\n"
fi


# Set ECC state
echo
echo "╭─────────────────────────────────╮"
echo "│   Changing ECC state for test   │"
echo "│                                 │"
echo "╰─────────────────────────────────╯"
echo
echo -e "Which test are you running?  (1)Passthrough    (2)vGPU\n"
read -p "Select an option: " OPTION
while [[ $OPTION != [12] ]]; do
    read -p "Select an option: " OPTION
done
ECC_state=`nvidia-smi -q | grep -EA2 'ECC Mode' | grep 'Current' | awk -F ': ' '{print $2}'`
case $OPTION in
"1")
    #  Enable ECC for Passthrough test
    echo
    echo "╭──────────────────────────────────────╮"
    echo "│   Enabling ECC for Passthrough test  │"
    echo "│                                      │"
    echo "╰──────────────────────────────────────╯"
    echo
    if [[ $ECC_state == 'Disabled' ]]; then
        nvidia-smi -e 1
        [[ $? == 0 ]] && systemctl reboot || { echo "❌ ERROR: Failed to enable ECC"; exit 1; }
    else
        echo -e "\n✅ ECC is enbled for Passthrough test"
    fi
    ;;
"2") 
    #  Disable ECC for vGPU test
    echo
    echo "╭────────────────────────────────╮"
    echo "│   Disabling ECC for vGPU test  │"
    echo "│                                │"
    echo "╰────────────────────────────────╯"
    echo
    if [[ $ECC_state == 'Enabled' ]]; then
        nvidia-smi -e 0
        [[ $? == 0 ]] && systemctl reboot || { echo "❌ ERROR: Failed to disable ECC"; exit 1; }
    else
        echo -e "\n✅ ECC is disbled for vGPU test"
    fi
    ;;
esac
[[ $? == 0 ]] && echo -e "\n\e[32mAll set! You are okay to go :)\e[0m\n"


exit

