@echo off
powershell -NoExit -Command "Set-ExecutionPolicy Bypass -Scope Process -Force; & '%~dp0arasaka-print-bridge.ps1' %*"
