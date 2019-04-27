# VBAstomp
A repository of example VBA stomped documents. The **.doc** and **.xls** extensions are files saved in the 97-2003 format, while the **docm** and **xlsm** file formats are save in the 2007 and up format (a zip archive xml format).

For more information about VBA Stomping, see [vbastomp.com](https://vbastomp.com) 

## Stomped Files Random vs Fake Code

Files tagged with **_random** have had their VBA source code replaced with random bytes using the [Adaptive Docoument Builder](https://github.com/haroldogden/adb) Python script.

Files tagged with **_fakecode** have had their VBA source code replaced with the source code found in **fakecode_word_vba.txt** or **fakecode_excel_vba.txt** using the [Evil Clippy](https://github.com/outflanknl/EvilClippy) tool.

## Filename Convention

The file name starts with the Office version and indicates if it is a 32-bit or 64-bit install.

Prefix|Description
---|---
2003x32|Office 2003, 32-bit
2019x64|Office 2019, 64-bit
