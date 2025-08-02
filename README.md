
# ReihenCheckXL

**ReihenCheckXL** ist ein Excel-Add-In zur Analyse von Zahlenreihen mit Fokus auf das Erkennen von Lücken. Ideal für Nutzer, die mit laufenden Nummern, Serien oder Datenreihen arbeiten und dabei fehlende Werte identifizieren möchten.

---

## 🛠️ Funktionen

### `LückenListe(Bereich)`
- Gibt eine Liste mit den fehlenden Zahlen in einer sortierten Reihe zurück.
- Beispiel: Aus der Reihe `1, 2, 3, 5, 6, 9` wird `4, 7, 8`

### `LückenArray(Bereich)`
- Gibt ein Array mit den fehlenden Zahlen zurück, das in Zellen aufgeteilt werden kann.
- Besonders nützlich für die Weiterverarbeitung in Excel-Formeln.

---

## 📥 Installation

1. Lade die Datei `ReihenCheckXL.xlam` herunter.
2. Öffne Excel, gehe zu **Datei > Optionen > Add-Ins**.
3. Wähle unten **Excel-Add-Ins** und klicke auf **Gehe zu...**
4. Klicke auf **Durchsuchen…** und wähle die Datei `ReihenCheckXL.xlam`.
5. Hake das Add-In an und bestätige – fertig!

---

## 🎯 Beispiel

Angenommen, du hast folgende Zahlenreihe in Excel:

```
A1:A6 = 1
        2
        3
        5
        6
        9
```

Dann liefert die Funktion:

```excel
=LückenListe(A1:A6)
```

den Text:  
`4, 7, 8`

Oder:

```excel
=LückenArray(A1:A6)
```

füllt die benachbarten Zellen automatisch mit:

```
B1 = 4
B2 = 7
B3 = 8
```

---

## 📄 Lizenz

Dieses Projekt verwendet die [MIT-Lizenz](https://choosealicense.com/licenses/mit/). Das bedeutet:
- Du kannst den Code frei nutzen, ändern und weitergeben.
- Bitte erwähne den ursprünglichen Autor.

---

## 📚 Weitere Dateien

- `src/VBA_Module.bas`: Exportierter VBA-Quelltext
- `examples/Testdatei.xlsx`: Beispieldatei mit einer Zahlenreihe zur Demonstration
- `docs/Anleitung.pdf`: Optionales PDF mit bebilderter Schritt-für-Schritt-Erklärung

---

## 🤝 Mitmachen

Feedback, Ideen oder Bug-Meldungen? Öffne gerne ein Issue oder starte einen Pull Request – jede Verbesserung macht das Projekt besser!


