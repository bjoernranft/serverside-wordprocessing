## Serverseitiges generieren von Dokumenten .NET Core / OpenXML-SDK

Das Projekt ist eine ASP.NET Core REST API und nutzt das MS-OpenXML-SDK Um serverseitig ein Dokument zu generieren 
Eine Office-Installation wird nicht benötigt.

GET-Request an die API. 

```
http://localhost:64457/api/values/Felix/WollMux
```

Der Aufruf erzeugt ein neues Word-Dokument, legt als Beispiel 
 - ein ContentControl mit Tag (=fullname) an
 - sucht das angelegte ContentControl und ersetzt den Text durch die Vorname/Nachname des GET-Requests.
 - erstellt und speichert einen Custom-Style (Formatvorlage)
 - erstellt Paragraph und wendet Custom-Style an.
 - Speichert .docx in hardcodierten Pfad "C:\\TestDoc\test.docx", bitte Verzeichnis anlegen falls nicht vorhanden.

