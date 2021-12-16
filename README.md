# StartMenu-to-Desktop

Creating Program Desktop Icons based on StartMenu entries (using Intune).


## Example: Creating Desktop Icon for LibreOffice Writer


### App-Typ:

__Windows-App (Win32)__

App-Paketdatei auswählen:

[./StartMenu-to-Desktop.intunewin](./StartMenu-to-Desktop.intunewin?raw=true)


### Name:

```
LibreOffice Writer Desktop Icon
```

### Description (Beschreibung):

```
Add LibreOffice Writer Desktop Icon for all users (system context)
```

### Publisher (Herausgeber)

```
BRG4-IT
```

### Install:
```
powershell.exe -executionpolicy bypass -file ".\StartMenu-to-Desktop.ps1" -AppName "LibreOffice Writer" -Label "Writer"
```


### Uninstall:
```
powershell.exe -executionpolicy bypass -file ".\StartMenu-to-Desktop.ps1" -AppName "LibreOffice Writer" -Label "Writer" -Remove
```

Install behaviour (Installationsverhalten): __System__


### Detection rules (Erkennungsregeln):

Regelformat (Rule type): __Erkennungsregeln manuell konfigurieren__

Rule type/Regel Typ: File/Datei

Path/Pfad:

```
%PUBLIC%\Desktop\
```


File or Folder/Datei oder Ordner:

```
Writer.lnk
```

Detection method: File or folder exists


### Dependencies (Abhängigkeiten):

LibreOffice
