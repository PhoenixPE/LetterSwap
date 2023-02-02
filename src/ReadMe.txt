LetterSwap.exe - позволяет переименовать буквы дисков на основе реестра гостевой системы.
Так-же возможно присвоить загрузочному диску букву Y: (ищется по маркерному файлу).
Синтаксис ком. строки :
LetterSwap.exe [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/BootDrive NewLetter:\DriveMarkerFile] [/log LogFile] [/Letter RegExp] [/wait 10]
/? - help
WinDir - прямое указание гостевой системы, например d:\windows
/Auto - находит первую гостевую систему
/Manual - выдает запрос
/bootdrive y:\MarkerFile присваивает диску с маркерным файлом букву Y:, удобно, если на нем есть программы, требующие абсолютный путь.
/HideLetter - скрывает буквы неактивных дисков
/MountAll - показывает все
/Log LogFile - создание лога

Пример:
LetterSwap.exe /HideLetter /auto /bootdrive y:\MarkerFile /Log "%temp%\LSLog.txt"

Утилита принудительно меняет диски, даже если они уже используются (буквы abxyz игнорируются), поэтому ее желательно запускать как можно раньше, например через RunOnceEx.
