Attribute VB_Name = "Modul1"


Sub Tagessumme_uebernehmen()

'Wenn Startgewicht < Zielgewicht  -->  Exclamationsmeldung und Exit
    If ((Worksheets("Meine Kalorientabelle").Range("B2").Value < Worksheets("Meine Kalorientabelle").Range("D2").Value)) Then

        Hallo = MsgBox("Diese Aufzeichnung hat das Ziel des Abnehmens", vbExclamation, "Meldung - byRichards Kalorientabelle f�r Chihauhas und andere Kleinhunde")

        Exit Sub

    End If

' Wenn Startgewicht = Zielgewicht oder Zielgewicht >= Tagesgewicht --> Gluckwunschsmeldung
    If ((Worksheets("Meine Kalorientabelle").Range("B2").Value = Worksheets("Meine Kalorientabelle").Range("D2").Value)) Or ((Worksheets("Meine Kalorientabelle").Range("D2").Value >= Worksheets("Meine Kalorientabelle").Range("H2").Value)) Then

        Hallo = MsgBox("Sie haben Ihr Ziel erreicht", vbInformation, "Gluckwunsch - byRichards Kalorientabelle f�r Chihauhas und andere Kleinhunde")



    End If



    Hierhingehen = Worksheets("Tagessummen").Cells(Rows.Count, 1).End(xlUp).Row


    If (Worksheets("Tagessummen").Range("A" & Hierhingehen).Value = Date) Then



'Worksheets("Tagessummen").Rows("hierhingehen").Delete Shift:=xlUp

        Worksheets("Tagessummen").Range("A" & Hierhingehen).EntireRow.Delete



        Worksheets("Tagessummen").Range("A" & Hierhingehen).Value = Date
        Worksheets("Tagessummen").Range("D" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("C5")
        Worksheets("Tagessummen").Range("E" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("D5")
        Worksheets("Tagessummen").Range("F" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("E5")
        Worksheets("Tagessummen").Range("G" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("F5")
        Worksheets("Tagessummen").Range("H" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("G5")
        Worksheets("Tagessummen").Range("I" & Hierhingehen).Value = Worksheets("Meine Kalorientabelle").Range("H2")

    Else

        Worksheets("Tagessummen").Range("A" & Hierhingehen + 1).Value = Date
        Worksheets("Tagessummen").Range("D" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("C5")
        Worksheets("Tagessummen").Range("E" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("D5")
        Worksheets("Tagessummen").Range("F" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("E5")
        Worksheets("Tagessummen").Range("G" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("F5")
        Worksheets("Tagessummen").Range("H" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("G5")
        Worksheets("Tagessummen").Range("I" & Hierhingehen + 1).Value = Worksheets("Meine Kalorientabelle").Range("H2")


    End If

'Worksheets("Tagessummen").UsedRange.RemoveDuplicates Columns:=1, Header:=xlYes



'Zusatzfragen Aufgaben


'Worksheets("Tagessummen").Select   ????? um zu sehen dass er Wirklich �bernommen hat



'Mit Startdatum und aktuelles Datum die restliche Zeit zum Zielgewicht berechen

'If (Worksheets("Meine Kalorientabelle").Range("B2").Value > Worksheets("Meine Kalorientabelle").Range("D2").Value) Then

'Hallo = MsgBox("Nur x Tage zum Ziel", VbInformation, "Meldung - byRichards Kalorientabelle f�r Chihauhas und andere Kleinhunde")



'End If


End Sub
