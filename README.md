traffic_generator
=================

Excel VBA Makro zur Erzeugung simulierter Recommendation Engine Aufrufe

## Voraussetzungen

Entwickelt wurde das Makro in Excel 2011 (Mac). Es sollte aber auch in den Windows Versionen ab 2007 funktionieren.


## Anwendung

* Produkt.xslm runterladen und öffnen
* den Makro Ausführen Dialog öffen (in Excel 2011 Menü: Extras/Makro/Makros...)
* das Makro "generateTraffic" ausführen
* alternativ mit Alt-F11 die Entwicklungsumgebung öffnen und das Makro direkt im Modul basDataGenerator  ausführen

## Ergebnis

Es wird ein neues Worksheet "Tesdaten<x>" eingefügt. Diese enthält für 8000 eindeutige Besucher Recommendation Aufrufe.
Dabei sieht sich jeder Besucher 1-5 Produkte an und jeder 30. Kunde bestellt etwas.
