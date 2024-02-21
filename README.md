# Restart Windows Explorer

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
2. **Run the Script**: Execute the script by entering the following command, `.\RestartExplorerRemotely.ps1` and when prompted, enter the target computer's name and credentials with administrative privileges on that computer.

### Disclaimer
This script is provided "as is", without warranty of any kind. Use it at your own risk. The author is not responsible for any damage or loss resulting from its use. Always test scripts in a non-production environment before deploying them to production.

### Contributing

Your contributions to improve the script or address issues are welcome. Please feel free to fork the repository, make your changes, and submit a pull request.
