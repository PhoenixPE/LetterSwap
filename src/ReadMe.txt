LetterSwap.exe - ��������� ������������� ����� ������ �� ������ ������� �������� �������.
���-�� �������� ��������� ������������ ����� ����� Y: (������ �� ���������� �����).
��������� ���. ������ :
LetterSwap.exe [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/BootDrive NewLetter:\DriveMarkerFile] [/log LogFile] [/Letter RegExp] [/wait 10]
/? - help
WinDir - ������ �������� �������� �������, �������� d:\windows
/Auto - ������� ������ �������� �������
/Manual - ������ ������
/bootdrive y:\MarkerFile ����������� ����� � ��������� ������ ����� Y:, ������, ���� �� ��� ���� ���������, ��������� ���������� ����.
/HideLetter - �������� ����� ���������� ������
/MountAll - ���������� ���
/Log LogFile - �������� ����

������:
LetterSwap.exe /HideLetter /auto /bootdrive y:\MarkerFile /Log "%temp%\LSLog.txt"

������� ������������� ������ �����, ���� ���� ��� ��� ������������ (����� abxyz ������������), ������� �� ���������� ��������� ��� ����� ������, �������� ����� RunOnceEx.
