# Maßlehre+ – GitHub Pages

Arbeitsstand für die browserbasierte Maßlehre+ Anwendung.

## Enthalten
- Responsive Weboberfläche mit Hamburger-Menü und Rechnernavigation
- Excel-Arbeitsblätter als JSON-Datenmodelle
- JavaScript-Formel-Laufzeit für die in der Arbeitsmappe verwendeten Funktionen
- Zellbasierte Eingabe- und Ausgabedarstellung
- Referenz- und Hilfsblätter bleiben als Daten erhalten

## Datenquelle
Datei: `SUEHAC_Maßlehre+_V2.1.6(9.2).xlsm`  
Interne Versionsangabe aus `Update Infos!B1`: **V2.1.9.3**  
Tabellenblätter: **54**  
Excel-Formeln: **746**

## GitHub Pages aktivieren
1. Repository auf GitHub anlegen.
2. Den Inhalt dieses Ordners in das Repository hochladen (also `index.html`, `css/`, `js/`, `data/` direkt ins Repository-Root).
3. In GitHub **Settings → Pages** öffnen.
4. `Deploy from a branch`, Branch `main`, Ordner `/ (root)` auswählen.
5. Nach dem Deployment die angezeigte GitHub-Pages-Adresse öffnen.

## WordPress-Domain
Die Anwendung kann später über eine Subdomain deiner WordPress-Domain erreichbar gemacht werden. Dafür wird im DNS ein CNAME-Eintrag auf die GitHub-Pages-Adresse gesetzt und die Custom Domain zusätzlich in den GitHub-Pages-Einstellungen hinterlegt. Die konkrete DNS-Maske hängt vom Domainanbieter ab.

## Fachlicher Prüfhinweis
Die Webanwendung berechnet dynamisch im Browser. Vor einem produktiven Einsatz müssen reale Eingabefälle pro Rechner mit Excel verglichen werden. Die enthaltenen Testdaten vergleichen Formelergebnisse, soweit Excel in der Quelldatei gespeicherte Ergebnis-Caches bereitgestellt hat. Insbesondere Eingabevarianten, Dropdowns und Grenzwerte benötigen zusätzliche fachliche Abnahme.


## Technischer Formeltest

Die JavaScript-Berechnungslogik wurde für 94 Formelergebnisse aus sechs Rechnerblättern mit neu berechneten LibreOffice-Referenzwerten verglichen. In diesem Test stimmten 94 von 94 Ergebnissen überein. Zusätzlich stimmen alle 676 Formeln mit vorhandenem Ergebnis-Cache im Ausgangsarbeitsblatt überein. Das ist ein technischer Regressionstest, ersetzt aber keine vollständige fachliche Abnahme sämtlicher Eingabekombinationen.


## Rechner- und Referenzblätter

Die Navigation enthält die Rechnerblätter sowie die zugehörigen Referenzblätter. `Inhaltsverzeichnis` und `Statistik` werden nicht als Rechnerseite angeboten. Referenzblätter wie `Zylinder`, `Classic-Glasmaße` oder `Bandabstand` bleiben als Tabellen erreichbar und werden nicht als Eingaberechner behandelt.


## Veröffentlichung: öffentlich, ohne Anmeldung

Der Nutzer hat bestätigt, dass die hochgeladene Arbeitsmappe die aktuellste und verbindliche Referenz ist. Die Maßlehre+ soll öffentlich und ohne Benutzeranmeldung nutzbar sein. Es gibt keine Login-Abfrage.

Die Weboberfläche und die benötigten JSON-Tabellenmodelle sind öffentlich abrufbar. Die originale `.xlsm`-Datei bleibt im getrennten privaten Repository und wird nicht auf GitHub Pages bereitgestellt.

Für die kostenlose Standardveröffentlichung mit GitHub Pages:
- Das Repository sollte öffentlich sein.
- Die Website wird über **Settings → Pages** aus `main` und `/ (root)` veröffentlicht.
- HTML-, CSS-, JavaScript- und JSON-Dateien sind öffentlich abrufbar.

## Eigene Domain verbinden

Eine eigene Domain oder Subdomain kann in GitHub unter **Settings → Pages → Custom domain** eingetragen werden. Danach muss beim DNS-Anbieter der passende DNS-Eintrag gesetzt werden. Der genaue Eintrag hängt davon ab, ob eine Hauptdomain oder Subdomain verwendet wird.

Eine WordPress-Website kann parallel bestehen bleiben. Die Maßlehre+ kann über eine eigene Subdomain wie `masslehre.deine-domain.de` auf GitHub Pages zeigen. WordPress wird dadurch nicht ersetzt.
