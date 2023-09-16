Attribute VB_Name = "Transform_Paie"

Sub Transform_Paie_Macro()

' cherche le matricule et le transforme en nom prenom et id
' Parcour chaque colonne d'une feuille 

File_Start = "Copie"
File_End = "Result"
Id_Ligne = 2            'ligne du tableau result
Id_Ligne_max = 0        'ligne max du tableau result
Col_Max = 0             'Col max du tableau Copie
Check=1                 'Visualise la fin du tableau

Sheets.Add.Name = File_End  'Ajout feuille

For Each col In Worksheets(File_Start).Columns

    i=i+1
    If i = 3  AND Check=1 Then
        Col_First = Split(Cells(, col.Column).Address, "$")(1)                              'Renvoie la lettre de la colonne : premiere colonne
        Matricule_all = Sheets(File_Start).Range(Col_First & "3").Value                     'Prend la valeur du matricule

        Col_Second = Split(Cells(, (col.Column+1)).Address, "$")(1)                         'Renvoie la lettre de la colonne : Deuxieme colonne
        Salaire_Base = Sheets(File_Start).Range(Col_Second & "7").Value                     'Prend la valeur du Salaire_Base
        Heures_Annualisees = Sheets(File_Start).Range(Col_Second & "19").Value              'Prend la valeur du Heures_Annualisees
        Absences = Sheets(File_Start).Range(Col_Second & "47").Value                        'Prend la valeur du Absences
        Mois = Sheets(File_Start).Range(Col_Second & "59").Value                            'Prend la valeur du Mois
        Sous_Total_Prime = Sheets(File_Start).Range(Col_Second & "128").Value               'Prend la valeur du Sous_Total_Prime
        Brut_SS = Sheets(File_Start).Range(Col_Second & "134").Value                        'Prend la valeur du Brute_SS
        Brut_Tranche_A = Sheets(File_Start).Range(Col_Second & "138").Value                 'Prend la valeur du Brut_Tranche_A
        Brut_Tranche_B = Sheets(File_Start).Range(Col_Second & "139").Value                 'Prend la valeur du Brut_Tranche_B
        Charge_Salariale = Sheets(File_Start).Range(Col_Second & "207").Value               'Prend la valeur du Charge_Salariale
        Retenue_Source = Sheets(File_Start).Range(Col_Second & "218").Value                 'Prend la valeur du Retenue_Source
        Prime_Pouvoir_Achat = Sheets(File_Start).Range(Col_Second & "219").Value            'Prend la valeur du Prime_Pouvoir_Achat
        Indemnite_Inflation = Sheets(File_Start).Range(Col_Second & "220").Value            'Prend la valeur du Indemnite_Inflation
        Net_Imposable = Sheets(File_Start).Range(Col_Second & "208").Value                  'Prend la valeur du Net_Imposable
        Net_Payer = Sheets(File_Start).Range(Col_Second & "260").Value                      'Prend la valeur du Net_Payer

        Col_Third = Split(Cells(, (col.Column+2)).Address, "$")(1)                          'Renvoie la lettre de la colonne : Troisieme colone
        Charge_Patronale = Sheets(File_Start).Range(Col_Third & "207").Value                'Prend la valeur du Charge_Patronale

        If Not Matricule_all = "TOTAL" Then
            Split_Matricule = Split(Matricule_all, " ")

            Sheets(File_End).Range("A" & Id_Ligne & ":C" & Id_Ligne) = Split_Matricule  'Matricule
            Sheets(File_End).Range("D" & Id_Ligne) = Salaire_Base                       'Salaire base
            Sheets(File_End).Range("E" & Id_Ligne) = Heures_Annualisees                 'Heures_Annualisees
            Sheets(File_End).Range("F" & Id_Ligne) = Absences                           'Absences
            Sheets(File_End).Range("G" & Id_Ligne) = Mois                               'Mois
            Sheets(File_End).Range("H" & Id_Ligne) = Sous_Total_Prime                   'Sous_Total_Prime
            Sheets(File_End).Range("J" & Id_Ligne) = Brut_SS                            'Brut_SS
            Sheets(File_End).Range("K" & Id_Ligne) = Brut_Tranche_A                     'Brut_Tranche_A
            Sheets(File_End).Range("L" & Id_Ligne) = Brut_Tranche_B                     'Brut_Tranche_B
            Sheets(File_End).Range("M" & Id_Ligne) = Charge_Patronale                   'Charge_Patronale
            Sheets(File_End).Range("N" & Id_Ligne) = Charge_Salariale                   'Charge_Salariale
            Sheets(File_End).Range("O" & Id_Ligne) = Retenue_Source                     'Retenue_Source
            Sheets(File_End).Range("P" & Id_Ligne) = Prime_Pouvoir_Achat                'Prime_Pouvoir_Achat
            Sheets(File_End).Range("Q" & Id_Ligne) = Indemnite_Inflation                'Indemnite_Inflation
            Sheets(File_End).Range("R" & Id_Ligne) = Net_Imposable                      'Net_Imposable
            Sheets(File_End).Range("S" & Id_Ligne) = Net_Payer                          'Net_Payer

            'Mise a 0 des champ vide
            If Mois = "" Then
                Mois=0
            End If
            If Sous_Total_Prime = "" Then
                Sous_Total_Prime=0
            End If
            If Salaire_Base = "" Then
                Salaire_Base=0
            End If
            If Charge_Salariale = "" Then
                Charge_Salariale=0
            End If
            'Mise a 0 des champ vide

            Sheets(File_End).Range("I" & Id_Ligne) = (Mois - Sous_Total_Prime)          
            Sheets(File_End).Range("T" & Id_Ligne) = (Salaire_Base + Sous_Total_Prime - Charge_Salariale)   
            
            Id_Ligne = Id_Ligne+1
        Else
            Id_Ligne_max=Id_Ligne
            Col_Max=col.Column
            Check=0
        End If
        i=0
    End If

Next col
MsgBox "Finish"
End Sub