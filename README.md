# Leankit-PoSH
A PowerShell module to interact with Leankit by Planview


To use this module you must load it and for the first run, create a configuration file by running the following command

```powershell
Import-Module -Name 'path\Leankit-PoSH.psd1' -Force
set-LKConfig -LKconfigFile <path\filename>
```
Once a config file is created you can initialize the module with the following command

```powershell
Initialize-LKToken -configFile <path\filename>