# VBAstomp

A repository of example VBA stomped documents. For more information about VBA Stomping, see [vbastomp.com](https://vbastomp.com) 

## Stomped Files Random vs Fake Code

Files tagged with **_random** have had their VBA source code replaced with random bytes using the [Adaptive Docoument Builder](https://github.com/haroldogden/adb) Python script.

Files tagged with **_fakecode** have had their VBA source code replaced with the source code found in **fakecode_word_vba.txt** or **fakecode_excel_vba.txt** using the [Evil Clippy](https://github.com/outflanknl/EvilClippy) tool.

## Filename Convention

The file name starts with the **Office version** and indicates if it is a 32-bit or 64-bit install.

Prefix|Description
---|---
2003x32|Office 2003, 32-bit
2019x64|Office 2019, 64-bit

The **Office version** is followed by either **_word** or **_excel** indicating whether it is an MS Word or MS Excel document.

The last part of the filename indicates whether the VBA source code was replace with random bytes or with new source code as described in the section above.

The **.doc** and **.xls** extensions are files saved in the 97-2003 format, while the **docm** and **xlsm** file formats are save in the 2007 and up format (a zip archive xml format).

## Original Files

The original files can be found in the **original_files_b4_stomping** folder.
