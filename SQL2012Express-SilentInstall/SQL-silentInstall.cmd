@echo off
pushd "%~dp0"
setlocal

set p=%~dp0

<SETUPSQL.EXE> /configurationfile="%p%ConfigurationFile.ini"

popd
endlocal