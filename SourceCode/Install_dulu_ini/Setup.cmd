regedit.exe /s IntegratedRegsvr.reg
regsvr32.exe THREED32.OCX
regedit.exe /s VBCTRLS.REG
regsvr32.exe XPControls.ocx
regsvr32.exe CRYSTL32.OCX
regsvr32.exe MSFLXGRD.OCX
regsvr32.exe TABCTL32.OCX
regsvr32.exe MSCOMCTL.OCX
regsvr32.exe mscomct2.ocx
copy THREED32.OCX %systemroot%\system32
copy XPCONTROLS.OCX %systemroot%\system32
copy CRYSTL32.OCX %systemroot%\system32
copy MSFLXGRD.OCX %systemroot%\system32
copy TABCTL32.OCX %systemroot%\system32
copy MSCOMCTL.OCX %systemroot%\system32
copy mscomct2.ocx %systemroot%\system32
copy ROMAN.FON %systemroot%\fonts