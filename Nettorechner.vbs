' Eingabe
brutto = InputBox("Geben Sie den Brutto-Betrag ein!","Eingabe Brutto")
' Verarbeitung (Berechnung)
netto = brutto / 119 * 100
'Ausgabe
ergebnis = "Brutto:" & vbTab & vbTab & FormatCurrency(brutto) & _
           vbNewline & _
           "/ 119 * 100:" & vbTab & FormatCurrency(netto) & _
           vbNewline & _
           "Gesamt: " & vbTab & vbTab & FormatCurrency(netto)


MsgBox ergebnis,,"Ergebnis"