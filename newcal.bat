call kv-activate
del vcconvert2.log
@echo on
python vcconvert2.py
@echo off
deactivate
copy stays.txt "g:\My Drive\VillaRaspi"
copy pool_heater_allowed.txt "g:\My Drive\VillaRaspi\pool_heater_allowed"
dir "g:\My Drive\VillaRaspi\stays.txt"
dir "g:\My Drive\VillaRaspi\pool_heater_allowed\pool_heater_allowed.txt"
rem eof
