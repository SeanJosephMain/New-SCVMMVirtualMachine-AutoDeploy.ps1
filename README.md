# New-SCVMMVirtualMachine-AutoDeploy.ps1
This PowerShell script automates VM creation in SCVMM. It checks cluster health, cleans up previous templates, selects optimal storage, and builds new VMs with defined CPU, memory, and network profiles. All paths, IDs, and webhooks are parameterized for security, with optional Slack alerts and detailed logging for each step.
