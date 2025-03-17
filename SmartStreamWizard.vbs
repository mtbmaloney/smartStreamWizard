If Not WScript.Arguments.Named.Exists("elevate") Then
CreateObject("Shell.Application").ShellExecute WScript.FullName _
, """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
WScript.Quit
End If

set shell = CreateObject("WScript.Shell")
set Network = CreateObject("WScript.Network")


'shell.Run("""c:\windows\system32\drivers\etc\hosts""")
'shell.Run("""c:\windows\syswow64\cliconfig.exe""")
'shell.Run("""c:\windows\syswow64\odbcad32.exe""")



set fso=CreateObject("scripting.fileSystemObject")
fso.copyfile "\\elect-apps\FINPROD\ntwdblib.dll", "C:\windows\system32\"

set file=fso.opentextfile("C:\windows\system32\drivers\etc\hosts")

content=file.ReadAll

'file=nothing
set writer=fso.opentextfile("C:\windows\system32\drivers\etc\hosts",8,Fasle)

set re= new RegExp
re.Pattern= "192.168.200.238"
if re.Test(content)=False then
  writer.write vbcrlf& "192.168.200.238    GOB0003"
end if

re.Pattern= "192.168.200.34"
if re.Test(content)=False then
  writer.write vbcrlf& "192.168.200.34    GOB0004"
end if

'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\FINPROD /v Database /t REG_EXPAND_SZ /d DBSctlg")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\FINPROD /v Driver /t REG_EXPAND_SZ /d C:\Windows\system32\SQLSRV32.dll")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\FINPROD /v Description /t REG_EXPAND_SZ /d GOB0003")

'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\HRPROD /v Database /t REG_EXPAND_SZ /d DBSctlg")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\HRPROD /v Driver /t REG_EXPAND_SZ /d C:\Windows\system32\SQLSRV32.dll")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\HRPROD /v Description /t REG_EXPAND_SZ /d GOB0003")

'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0003 /v Database /t REG_EXPAND_SZ /d DBSctlg")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0003 /v Driver /t REG_EXPAND_SZ /d C:\Windows\system32\SQLSRV32.dll")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0003 /v Description /t REG_EXPAND_SZ /d GOB0003")

'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0004 /v Database /t REG_EXPAND_SZ /d DBSctlg")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0004 /v Driver /t REG_EXPAND_SZ /d 1C:\Windows\system32\SQLSRV32.dll")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\GOB0004 /v Description /t REG_EXPAND_SZ /d GOB0004")

'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\DBSzcrd /v Database /t REG_EXPAND_SZ /d DBSzcrd")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\DBSzcrd /v Driver /t REG_EXPAND_SZ /d C:\Windows\system32\SQLSRV32.dll")
'shell.Run("reg add HKLM\SOFTWARE\ODBC\ODBC.INI\DBSzcrd /v Description /t REG_EXPAND_SZ /d GOB0003")

shell.Run("ODBCCONF.EXE /a {CONFIGSYSDSN ""SQL Server"" ""DSN=FINPROD|Description=GOB0003|SERVER=GOB0003|Trusted_Connection=Yes|Database=DBSctlg""}")
shell.Run("ODBCCONF.EXE /a {CONFIGSYSDSN ""SQL Server"" ""DSN=HRPROD|Description=GOB0004|SERVER=GOB0004|Trusted_Connection=Yes|Database=DBSctlg""}")
shell.Run("ODBCCONF.EXE /a {CONFIGSYSDSN ""SQL Server"" ""DSN=GOB0003|Description=GOB0003|SERVER=GOB0003|Trusted_Connection=Yes|Database=DBSctlg""}")
shell.Run("ODBCCONF.EXE /a {CONFIGSYSDSN ""SQL Server"" ""DSN=GOB0004|Description=GOB0004|SERVER=GOB0004|Trusted_Connection=Yes|Database=DBSctlg""}")
shell.Run("ODBCCONF.EXE /a {CONFIGSYSDSN ""SQL Server"" ""DSN=DBSzcrd|Description=DBSzcrd|SERVER=GOB0003|Trusted_Connection=Yes|Database=DBSzcrd""}")

Network.MapNetworkDrive "Q:" , "\\192.168.21.12\FINPROD\" 
shell.Run("""Q:\ss\setup.exe""")
