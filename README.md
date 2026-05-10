\# Stundenplan V14 – Automatischer Stundenplangenerator



Automatische Stundenplanerstellung für Gymnasien und weiterführende Schulen,

entwickelt mit C#, WPF und Google OR-Tools.



\## Was kann das Programm?



\- Vollautomatische Stundenplanerstellung per Constraint-Solver (Google OR-Tools)

\- Berücksichtigt Lehrer- und Klassenkonflikte, Zeitwünsche und Sperrzeiten

\- Doppelstunden, Fachraum-Limits und Lehrertauschgruppen

\- Mehrere Lösungen werden berechnet und nach Qualität bewertet

\- Iterative Planverbesserung

\- Export von Lehrer- und Klassenplänen als Excel-Sheets

\- Vollständige Constraint-Prüfung mit farbcodiertem Bericht

\- Bedienung über einfache Excel-Eingabedatei (.xlsx)



\## Voraussetzungen



\- Windows 10 oder höher

\- Microsoft Excel oder ein kompatibles Programm (zum Öffnen der Ergebnisse)

\- .NET 8 Runtime



\## Installation



1\. Unter \[Releases](https://github.com/Ansgar13/StundenplanV14/releases) 

&#x20;  die aktuelle `Stundenplan\_V14.exe` herunterladen

2\. EXE starten – keine Installation nötig



\## Bedienung



Eine ausführliche Gebrauchsanleitung (inkl. Excel-Aufbau, alle Buttons,

Makros und Arbeitsablauf) steht als Word-Dokument bereit:



📄 \[Stundenplan\_V14\_Anleitung.docx](Stundenplan\_V14\_Anleitung.docx)



\## Kurzanleitung



1\. Excel-Datei mit Unterrichtsdaten befüllen (Sheet „U-Verteilung")

2\. Programm starten → Button 2: Excel laden

3\. Button 3: Stundenplan erstellen

4\. Button 4/5: Lehrer- und Klassenpläne exportieren



\## Technologie



\- C# / WPF (.NET 8)

\- Google OR-Tools (Constraint Programming)

\- ClosedXML (Excel-Integration)



\## Lizenz



Frei verwendbar für Schulen und Bildungseinrichtungen.

