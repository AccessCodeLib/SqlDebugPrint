# SqlDebugPrint
Dieses AddIn erleichtert das Testen von in VBA zusammengesetzten (SQL-Strings/Kriterien-
ausdrücken), indem es die in Foren oft ans Herz gelegte Vorgehensweise:

 - String mit "Debug.Print string" ins Direktfenster schreiben
 - SQL-String aus dem Direktfenster kopieren
 - leere Abfrage erstellen
 - SQL-String in die Abfrage kopieren
 - Abfrage testen

auf ein Formular "umbiegt", wo mit Buttonklicks die Abfrage auf Fehler (werden in einem 
Textfeld angezeigt) überprüft, und eine Abfrage (temporär) in Datenblatt- bzw. SQL-Ansicht 
geöffnet werden kann.

### Installation
 - "SqlDebugPrint.zip" in ein beliebiges Verzeichnis entpacken
 - das Script "Install.vbs" ausführen, kopiert die Datei, auf Wunsch kompiliert (empfohlen)
   in das Access-AddIn-Verzeichnis
 - in einer beliebigen DB den AddIn-Manager aufrufen, das AddIn auswählen und installieren
 - AddIn einmal im Menu aufrufen
 - das Startformular erklärt alles Weitere
