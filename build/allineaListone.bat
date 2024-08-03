@echo off

echo "Allinea Listone Portieri..."
start /wait AllineaListoneExecutable.exe p
echo "Allinea Listone Difensori..."
start /wait AllineaListoneExecutable.exe d
echo "Allinea Listone Centrocampisti..."
start /wait AllineaListoneExecutable.exe c
echo "Allinea Listone Attaccanti..."
start /wait AllineaListoneExecutable.exe a

pause