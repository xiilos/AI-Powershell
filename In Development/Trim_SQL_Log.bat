@echo off
sqlcmd -S servername\instancename -d <DatabaseName> -E -Q "DBCC SHRINKFILE('A2E_log', 1);"




