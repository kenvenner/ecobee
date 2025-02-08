copy ..\poolcontrol\pool.py .
call kv-activate
del vcconvert2.log
@echo on
python vcconvert2.py
@echo off
deactivate
rem eof
