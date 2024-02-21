# Restart Windows Explorer Remotely Without PSRemoting

## Introduction

This PowerShell script enables IT administrators to remotely restart the Windows Explorer process on a target machine using CIM sessions instead of PSRemoting. This approach can be beneficial in environments where PSRemoting is not enabled by default or for quick troubleshooting on remote computers.

## Prerequisites

- PowerShell 5.0 or higher.
- Administrative privileges on the local machine.
- The target machine must allow DCOM communication, which is used by CIM sessions.
- Credentials with administrative privileges on the target machine.

## Getting Started

Before using the script, ensure that the target machine's firewall and security settings allow DCOM communication, which is essential for creating a CIM session.

### Enable DCOM Communication

1. On the target machine, open the Component Services administrative tool.
2. Navigate to Component Services > Computers > My Computer and right-click 'My Computer', then select Properties.
3. In the Properties dialog, go to the COM Security tab and configure the necessary permissions for Launch and Activation Permissions and Access Permissions.

### Usage Instructions

1. **Open PowerShell**: Launch PowerShell with administrative privileges on your local machine.
2. **Run the Script**: Execute the script by entering the following command, and when prompted, enter the target computer's name and credentials with administrative privileges on that computer.

```powershell
.\RestartExplorerRemotely.ps1

You will be prompted to enter the computer name and credentials for the target machine. The script will then attempt to enable PSRemoting on the target machine using a CIM session and proceed to restart the Windows Explorer process.

How It Works

The script performs the following actions:

Prompts the user for the target computer name and credentials.
Creates a CIM session using DCOM protocol to communicate with the target machine.
Tries to enable PSRemoting on the target machine as a prerequisite for further actions (this is an optional step and can be modified according to your needs).
Uses Invoke-Command to execute a script block that stops and then restarts the Windows Explorer process on the target machine.
Troubleshooting

If you encounter issues with DCOM communication or permissions, ensure that:

The target machine's firewall allows DCOM traffic.
The user account has the necessary permissions to perform DCOM operations on the target machine.
Disclaimer

This script is provided "as is", without warranty of any kind. Use it at your own risk. The author is not responsible for any damage or loss resulting from its use. Always test scripts in a non-production environment before deploying them to production.

Contributing

Your contributions to improve the script or address issues are welcome. Please feel free to fork the repository, make your changes, and submit a pull request.
```

