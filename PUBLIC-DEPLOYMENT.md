# Öffentliche Veröffentlichung der Maßlehre+

Die Maßlehre+ soll ohne Benutzerkonto und ohne Login öffentlich nutzbar sein.

## GitHub
1. ZIP entpacken.
2. Ein öffentliches GitHub-Repository mit dem Namen `masslehre-plus-pages` verwenden.
3. Den Inhalt des entpackten Projektordners direkt in das Repository-Root hochladen.
4. In GitHub `Settings → Pages` öffnen.
5. Als Quelle `Deploy from a branch`, Branch `main` und Ordner `/ (root)` auswählen.
6. Die angezeigte GitHub-Pages-Adresse aufrufen und Rechner auf Desktop und Smartphone testen.

## Eigene Domain
1. Gewünschte Domain/Subdomain unter `Settings → Pages → Custom domain` eintragen.
2. Den von GitHub erwarteten DNS-Eintrag beim Domainanbieter setzen.
3. Warten, bis GitHub die DNS-Konfiguration erkannt und HTTPS aktiviert hat.

Der genaue DNS-Eintrag hängt davon ab, ob eine Hauptdomain (z. B. `example.de`) oder eine Subdomain (z. B. `masslehre.example.de`) verwendet wird. Bestehende WordPress-DNS-Einträge nicht ungeprüft überschreiben.

## Hinweis
Die Anwendung ist öffentlich. Alle im Repository abgelegten Dateien, insbesondere JSON-Datenmodelle, können von Besuchern abgerufen werden. Keine vertraulichen oder personenbezogenen Daten einpflegen.
