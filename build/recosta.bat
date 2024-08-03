@echo off

echo "Portieri Recosta..."
start /wait /B RecostaExecutable.exe p
echo "Difensori Recosta..."
start /wait /B RecostaExecutable.exe d
echo "Centrocampisti Recosta..."
start /wait /B RecostaExecutable.exe c
echo "Attaccanti Recosta..."
start /wait /B RecostaExecutable.exe a

pause
