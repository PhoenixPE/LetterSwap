LetterSwap.exe - позволяет переименовать буквы дисков на основе реестра гостевой системы.
Так-же возможно присвоить загрузочному диску определенную букву : (ищется по маркерному файлу).
Синтаксис ком. строки :
LetterSwap.exe  [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/Save] [/BootDrive NewLetter:[\TagFile]] [/SetLetter NewLetter:\TagFile] [/RestartExplorer] [/log [LogFile|con:]] [/IgnoreLetter Letters] [/Swap Drive: Drive:] [/wait 10]
/? - help
WinDir - прямое указание гостевой системы, например d:\windows , можно использовать переменные среды, например  %OFFINESYSTEM%
/Auto - находит первую гостевую систему
/Manual - выдает запрос
/bootdrive y:\TagFile - присваивает диску с маркерным файлом букву Y:, удобно, если на нем есть программы, требующие абсолютный путь.
/bootdrive y: - присваивает загрузочному диску букву Y: , TagFile не требуется.
/SetLetter z:\File.tag  
/HideLetter - скрывает буквы неактивных дисков
/MountAll - показывает все
/Swap Drive: Drive: - поменять буквы
/Save - создать ветку реестра HKLM\SOFTWARE\LetterSwap,Guest,REG_SZ,БукваГостевойСистемы:
/Log LogFile - создание лога
/Letter RegExp - не трогать RegExp буквы, по умолчанию yz, SystemDrive и a: b: - всегда игнорируются.
/wait время ожидания монтирования BootDrive

Пример:
LetterSwap.exe /HideLetter /auto /bootdrive y: /Log "%temp%\LSLog.txt"

Утилита принудительно меняет диски, даже если они уже используются (буквы abxyz игнорируются), поэтому ее желательно запускать как можно раньше, например через RunOnceEx.

LetterSwap.a3x требует поддержки в системе Autoit версии  3.3.14.2