Option Explicit On
Option Infer Off
Option Strict On
'Main Form.vb
'TP1 - Application qui répartie les dépenses.
'2020-01-31
'111245796

Imports System.ComponentModel

Public Class frmMain

    Private Const decTPS As Decimal = 0.05D
    Private Const decTVQ As Decimal = 0.0975D

    Private decPourboire As Decimal
    Private decTotalAPayer As Decimal
    Private decTotalPaye As Decimal
    Private decRestant As Decimal
    Private decSommeRestaurant As Decimal
    Private decSommeTransport As Decimal
    Private decSommeHebergement As Decimal
    Private decSommeAutre As Decimal
    Private decSommeDepenses As Decimal
    Private decPayeTotalP1 As Decimal
    Private decPayeTotalP2 As Decimal
    Private decPayeTotalP3 As Decimal
    Private decPayeTotalP4 As Decimal

    ' Procédure Sub et Function indépendantes.
    Private Sub CalculerPourboire()
        ' Fais le calcul du pourboire selon le type.

        Const decPOURBOIRE_RESTAURANT As Decimal = 0.15D
        Const decPOURBOIRE_TRANSPORT As Decimal = 0.1D
        Const decPOURBOIRE_HEBERGEMENT As Decimal = 5
        Const decPOURBOIRE_AUTRES As Decimal = 0.15D

        Dim decTotalFacture As Decimal
        Dim decTotalAvantTaxes As Decimal

        ' Fais la convertion et calcul le montant avant taxe.
        Decimal.TryParse(txtTotalFacture.Text, decTotalFacture)
        decTotalAvantTaxes = decTotalFacture * (1 - (decTPS + decTVQ))

        ' Calcul le pouboire selon le type de dépense.
        Select Case DepenseType()
            Case 1
                ' Pourboire restaurant
                decPourboire = decTotalAvantTaxes * decPOURBOIRE_RESTAURANT
            Case 2
                ' Pourboire transport
                decPourboire = decTotalAvantTaxes * decPOURBOIRE_TRANSPORT
            Case 3
                ' Pourboire hébergement
                decPourboire = decPOURBOIRE_HEBERGEMENT
            Case 4
                ' Pourboire autre                
                decPourboire = decTotalAvantTaxes * decPOURBOIRE_AUTRES
        End Select

        lblPourboire.Text = decPourboire.ToString("C2")

    End Sub

    Private Sub CaractereTextBox(ByVal strNomTextBox As String, e As KeyPressEventArgs)
        ' S'assure qu'on utilise seulement les caractères 0 à 9, la virgule(,) et la touche Backspace.

        ' Si la virgule est déjà utilisé, elle ne peut pas être tapé une seconde fois.
        If ContientVirgule(strNomTextBox) Then
            If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If
        Else
            If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso
                e.KeyChar <> ControlChars.Back AndAlso
                e.KeyChar <> "," Then
                e.Handled = True
            End If
        End If

    End Sub

    Function DepenseType() As Integer
        ' Permet de savoir quel type de dépense a été sélectionné.
        ' Retour: Un nombre Integer, Restaurant = 1, Transport = 2, Hebergement = 3, Autre = 4
        '         Si aucune sélection = 0

        ' Selon le bouton radio coché.
        Select Case True
            Case radRestaurant.Checked
                Return 1
            Case radTransport.Checked
                Return 2
            Case radHebergement.Checked
                Return 3
            Case radAutre.Checked
                Return 4
            Case Else
                Return 0
        End Select

    End Function

    Function ContientVirgule(ByVal strTexte As String) As Boolean
        ' Valide si une variable de type String contient une virgule.
        ' Retour: Un Boolean True si il trouve une virgule, False sinon.

        Dim blnEstVirgule As Boolean

        ' Si le texte contient déjà une virgule.
        blnEstVirgule = strTexte.Contains(",")

        Return blnEstVirgule

    End Function

    Private Sub CalculMontantTotalPaye()
        ' Calcul le montant total payé par chaque personne pour la dépense courante.

        Dim decPayeP1 As Decimal
        Dim decPayeP2 As Decimal
        Dim decPayeP3 As Decimal
        Dim decPayeP4 As Decimal

        Decimal.TryParse(txtPayeP1.Text, decPayeP1)
        Decimal.TryParse(txtPayeP2.Text, decPayeP2)
        Decimal.TryParse(txtPayeP3.Text, decPayeP3)
        Decimal.TryParse(txtPayeP4.Text, decPayeP4)

        decTotalPaye = decPayeP1 + decPayeP2 + decPayeP3 + decPayeP4
        lblTotalPaye.Text = decTotalPaye.ToString("C2")

        'Met à jour le montant restant.
        CalculMontantRestant()

    End Sub

    Private Sub CalculMontantTotalAPayer()
        ' Calcul le montant total de la facture + le montant du pourboire de la dépense courante.

        Dim decTotalFacture As Decimal
        Dim blnIsTotalFactureValid As Boolean

        blnIsTotalFactureValid = Decimal.TryParse(txtTotalFacture.Text, decTotalFacture)

        ' Calcul le montant total avec le pourboire s'il y a lieu.
        If chkPourboire.Checked Then
            CalculerPourboire()
            decTotalAPayer = decTotalFacture + decPourboire
        Else
            decTotalAPayer = decTotalFacture
        End If

        lblTotalAPayer.Text = decTotalAPayer.ToString("C2")

        ' Met à jour le montant restant.
        CalculMontantRestant()

    End Sub

    Private Sub CalculMontantRestant()
        ' Calcul et le montant restant à l'aide du montant total payé moins le montant du total à payer.

        decRestant = decTotalPaye - decTotalAPayer

        'Affiche le montant dans le label Montant Restant.
        lblMntRestant.Text = decRestant.ToString("C2")

    End Sub

    Private Sub VideFormulaire()
        ' Vide le formulaire et remet les attributs à zéro.

        decPourboire = 0
        decTotalAPayer = 0
        decTotalPaye = 0
        decRestant = 0

        txtPayeP1.Text = String.Empty
        txtPayeP2.Text = String.Empty
        txtPayeP3.Text = String.Empty
        txtPayeP4.Text = String.Empty
        txtTotalFacture.Text = String.Empty
        chkPourboire.Checked = False
        lblPourboire.Text = String.Empty
        radRestaurant.Checked = True

        ' Réinitialise les calculs avec les champs vides.
        CalculMontantTotalAPayer()

    End Sub

    Private Sub CompteurNbDepenses()
        ' Compte le nombre de facture qui a été traité.

        Static intCompteurNbDepenses As Integer

        ' Incrémente le compteur de dépenses de 1.
        intCompteurNbDepenses += 1

        ' Met à jour le label du nombre de facture traité.
        lblNbDepenses.Text = intCompteurNbDepenses.ToString()

    End Sub

    Private Sub BloquerChangementPersonne()
        ' Empêche le changement de nom lorsqu'il y a eu au moins une entrée de dépense.

        ' Valide si la personne 2 est coché et bloque s'il y a lieu.
        If chkP2.Checked Then
            chkP2.Enabled = False
            lblNomP2.Enabled = False
            txtNomP2.Enabled = False
        Else
            chkP2.Enabled = False
        End If

        ' Valide si la personne 3 est coché et bloque s'il y a lieu.
        If chkP3.Checked Then
            chkP3.Enabled = False
            lblNomP3.Enabled = False
            txtNomP3.Enabled = False
        Else
            chkP3.Enabled = False
        End If

        ' Valide si la personne 4 est coché et bloque s'il y a lieu.
        If chkP4.Checked Then
            chkP4.Enabled = False
            lblNomP4.Enabled = False
            txtNomP4.Enabled = False
        Else
            chkP4.Enabled = False
        End If

    End Sub

    Function NombrePersonneActive() As Integer
        ' Compte le nombre de personne qui a été activé. 
        ' Retour: Un Integer avec le nombre de personne.

        Dim intCompteurPersonne As Integer = 1

        ' Si la personne 2 est coché, incrémente le compteur de un.
        If chkP2.Checked Then
            intCompteurPersonne += 1
        End If

        ' Si la personne 3 est coché, incrémente le compteur de un.
        If chkP3.Checked Then
            intCompteurPersonne += 1
        End If

        ' Si la personne 4 est coché, incrémente le compteur de un.
        If chkP4.Checked Then
            intCompteurPersonne += 1
        End If

        Return intCompteurPersonne

    End Function

    Private Sub RepartieEtAfficheMontant()
        ' Réparti et affiche les montants à recevoir.

        Dim decPayeP1 As Decimal
        Dim decPayeP2 As Decimal
        Dim decPayeP3 As Decimal
        Dim decPayeP4 As Decimal
        Dim decRecevoirP1 As Decimal
        Dim decRecevoirP2 As Decimal
        Dim decRecevoirP3 As Decimal
        Dim decRecevoirP4 As Decimal

        ' Calcule le montant à recevoir pour la personne 1 s'il y a lieu.
        Decimal.TryParse(txtPayeP1.Text, decPayeP1)
        decPayeTotalP1 += decPayeP1
        decRecevoirP1 = MntRepartieParPersonne(decPayeTotalP1)
        lblRecevoirP1.Text = decRecevoirP1.ToString("C2")

        ' Calcule le montant à recevoir pour la personne 2 s'il y a lieu.
        If chkP2.Checked Then
            Decimal.TryParse(txtPayeP2.Text, decPayeP2)
            decPayeTotalP2 += decPayeP2
            decRecevoirP2 = MntRepartieParPersonne(decPayeTotalP2)
            lblRecevoirP2.Text = decRecevoirP2.ToString("C2")
        End If

        ' Calcule le montant à recevoir pour la personne 3 s'il y a lieu.
        If chkP3.Checked Then
            Decimal.TryParse(txtPayeP3.Text, decPayeP3)
            decPayeTotalP3 += decPayeP3
            decRecevoirP3 = MntRepartieParPersonne(decPayeTotalP3)
            lblRecevoirP3.Text = decRecevoirP3.ToString("C2")
        End If

        ' Calcule le montant à recevoir pour la personne 4 s'il y a lieu.
        If chkP4.Checked Then
            Decimal.TryParse(txtPayeP4.Text, decPayeP4)
            decPayeTotalP4 += decPayeP4
            decRecevoirP4 = MntRepartieParPersonne(decPayeTotalP4)
            lblRecevoirP4.Text = decRecevoirP4.ToString("C2")
        End If

    End Sub

    Function BilanEcran() As String
        ' Format et présente un bilan en texte qui sera visible à l'écran.
        ' Retour: Un String formaté avec le bilan.

        Dim intNbPersonne As Integer = NombrePersonneActive()
        Dim strBilanFormate As String

        ' Affiche le nombre de dépenses, la répartition et la somme par catégorie.
        strBilanFormate = "Vous avez " & lblNbDepenses.Text & " dépenses qui ont été réparties entre " &
            intNbPersonne.ToString & " personnes." & ControlChars.NewLine & "Dépenses totales de " &
            decSommeDepenses.ToString("C2") & ControlChars.NewLine &
            decSommeRestaurant.ToString("C2") & " pour les restaurants." & ControlChars.NewLine &
            decSommeTransport.ToString("C2") & " pour les transports." & ControlChars.NewLine &
            decSommeHebergement.ToString("C2") & " pour l'hébergement." & ControlChars.NewLine &
            decSommeAutre.ToString("C2") & " pour les autres dépenses."

        ' Ajoute les personnes entre 1 et 4 qui doivent recevoir ou donner de l'argent, s'il y a lieu.
        If intNbPersonne > 1 Then
            strBilanFormate = strBilanFormate & ControlChars.NewLine & lblNomP1.Text & " doit " &
                RecevoirOuDonner(MntRepartieParPersonne(decPayeTotalP1)) & " " &
                AbsToString(MntRepartieParPersonne(decPayeTotalP1)) & "$" & ControlChars.NewLine & txtNomP2.Text &
                " doit " & RecevoirOuDonner(MntRepartieParPersonne(decPayeTotalP2)) & " " &
                AbsToString(MntRepartieParPersonne(decPayeTotalP2)) & "$"
            If intNbPersonne > 2 Then
                strBilanFormate = strBilanFormate & ControlChars.NewLine & txtNomP3.Text & " doit " &
                    RecevoirOuDonner(MntRepartieParPersonne(decPayeTotalP3)) & " " &
                    AbsToString(MntRepartieParPersonne(decPayeTotalP3)) & "$"
            End If
            If intNbPersonne > 3 Then
                strBilanFormate = strBilanFormate & ControlChars.NewLine & txtNomP4.Text & " doit " &
                    RecevoirOuDonner(MntRepartieParPersonne(decPayeTotalP4)) & " " &
                    AbsToString(MntRepartieParPersonne(decPayeTotalP4)) & "$"
            End If
        End If

        ' Ajoute les remerciements.
        strBilanFormate = strBilanFormate + ControlChars.NewLine + "Merci d'avoir utilisé Acomba Pro 95 et bonne journée!"

        Return strBilanFormate

    End Function

    Function RecevoirOuDonner(ByVal montant As Decimal) As String
        ' Détermine si le montant est à recevoir ou à donner.
        ' Retour: Une variable de type String contenant le mot "recevoir" ou "donner".

        Dim strRetour As String

        ' Si le montant est plus grand ou égale à zéro, la personne reçoit. Sinon la personne donne.
        If montant >= 0 Then
            strRetour = "recevoir"
        Else
            strRetour = "donner"
        End If

        Return strRetour

    End Function

    Function AbsToString(ByVal montant As Decimal) As String
        ' Fait la conversion vers un nombre absolu.
        ' Retour: Un String contenant la valeur absolue du montant.

        Dim decValeurAbs As Decimal

        ' Valeur absolue.
        decValeurAbs = Math.Abs(montant)

        Return decValeurAbs.ToString("C2")

    End Function

    Function MntRepartieParPersonne(ByVal montant As Decimal) As Decimal
        ' Reparti le montant calculé pour une personne.
        ' Retour: Un nombre Decimal contenant le montant réparti pour une personne.

        Dim decMontantSepare As Decimal
        Dim decRecevoir As Decimal

        ' Montant réparti par le nombre de personne.
        decMontantSepare = decSommeDepenses / NombrePersonneActive()

        decRecevoir = (decMontantSepare - montant) * -1

        Return decRecevoir

    End Function

    Private Sub btnQuitter_Click(sender As Object, e As EventArgs) Handles btnQuitter.Click

        Me.Close()

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Demande à l'utilisateur d'entré son nom à l'ouverture.

        Dim strNomUtilisateur As String

        strNomUtilisateur = InputBox("Entrer votre nom complet :", "Nom Utilisateur")

        ' Si le nom est vide, l'application ferme.
        If strNomUtilisateur.Trim = "" Then
            Me.Close()
        End If

        ' Écrit le nom de l'utilisateur dans le label de la personne 1.
        lblNomP1.Text = strNomUtilisateur

    End Sub

    Private Sub chkP2_Click(sender As Object, e As EventArgs) Handles chkP2.Click
        ' Active ou désactive la personne 2.

        ' Si coché active les champs, sinon désactivé.
        If chkP2.Checked Then
            lblNomP2.Enabled = True
            txtNomP2.Enabled = True
            txtPayeP2.Enabled = True
            lblDollarP2.Enabled = True
        Else
            lblNomP2.Enabled = False
            txtNomP2.Enabled = False
            txtNomP2.Text = String.Empty
            txtPayeP2.Enabled = False
            txtPayeP2.Text = String.Empty
            lblDollarP2.Enabled = False
        End If

    End Sub

    Private Sub chkP3_Click(sender As Object, e As EventArgs) Handles chkP3.Click
        ' Active ou désactive la personne 3.

        ' Si coché active les champs, sinon désactivé.
        If chkP3.Checked Then
            lblNomP3.Enabled = True
            txtNomP3.Enabled = True
            txtPayeP3.Enabled = True
            lblDollarP3.Enabled = True
        Else
            lblNomP3.Enabled = False
            txtNomP3.Enabled = False
            txtNomP3.Text = String.Empty
            txtPayeP3.Enabled = False
            txtPayeP3.Text = String.Empty
            lblDollarP3.Enabled = False
        End If

    End Sub

    Private Sub chkP4_Click(sender As Object, e As EventArgs) Handles chkP4.Click
        ' Active ou désactive la personne 4.

        ' Si coché active les champs, sinon désactivé.
        If chkP4.Checked Then
            lblNomP4.Enabled = True
            txtNomP4.Enabled = True
            txtPayeP4.Enabled = True
            lblDollarP4.Enabled = True
        Else
            lblNomP4.Enabled = False
            txtNomP4.Enabled = False
            txtNomP4.Text = String.Empty
            txtPayeP4.Enabled = False
            txtPayeP4.Text = String.Empty
            lblDollarP4.Enabled = False
        End If

    End Sub

    Private Sub chkPourboire_Click(sender As Object, e As EventArgs) Handles chkPourboire.Click
        ' Calcul le pourboire lorsque activé ou remet à zéro lorsque désactivé.

        ' Si coché active l'étiquette, sinon désactivé.
        If chkPourboire.Checked Then
            lblPourboireEtiquette.Enabled = True
        Else
            lblPourboire.Text = String.Empty
            lblPourboireEtiquette.Enabled = False
            decPourboire = 0
        End If

        CalculMontantTotalAPayer()

    End Sub

    Private Sub lblMntRestant_TextChanged(sender As Object, e As EventArgs) Handles lblMntRestant.TextChanged
        ' Lorsque le texte change, active ou désactive le bouton Répartir.

        ' Si le montant restant égale zéro et que le total à payer n'est pas égale à zéro.
        If decRestant > -0.01 AndAlso decRestant < 0.01 AndAlso decTotalAPayer <> 0 Then
            btnRepartir.Enabled = True
        Else
            btnRepartir.Enabled = False
        End If

    End Sub

    Private Sub txtPayeP1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPayeP1.KeyPress
        ' Valide les bon caractères.

        CaractereTextBox(txtPayeP1.Text, e)

    End Sub

    Private Sub txtPayeP2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPayeP2.KeyPress
        ' Valide les bon caractères.
        CaractereTextBox(txtPayeP2.Text, e)

    End Sub

    Private Sub txtPayeP3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPayeP3.KeyPress
        ' Valide les bon caractères.
        CaractereTextBox(txtPayeP3.Text, e)

    End Sub

    Private Sub txtPayeP4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPayeP4.KeyPress
        ' Valide les bon caractères.
        CaractereTextBox(txtPayeP4.Text, e)

    End Sub

    Private Sub txtTotalFacture_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTotalFacture.KeyPress
        ' Valide les bon caractères.
        CaractereTextBox(txtTotalFacture.Text, e)

    End Sub

    Private Sub txtTotalFacture_TextChanged(sender As Object, e As EventArgs) Handles txtTotalFacture.TextChanged
        ' Met à jour les calculs.

        CalculMontantTotalAPayer()

    End Sub

    Private Sub radAutre_CheckedChanged(sender As Object, e As EventArgs) Handles radAutre.CheckedChanged
        ' Met à jour le pourboire s'il y a lieu.

        CalculMontantTotalAPayer()

    End Sub

    Private Sub radHebergement_CheckedChanged(sender As Object, e As EventArgs) Handles radHebergement.CheckedChanged
        ' Met à jour le pourboire s'il y a lieu.

        CalculMontantTotalAPayer()

    End Sub

    Private Sub radRestaurant_CheckedChanged(sender As Object, e As EventArgs) Handles radRestaurant.CheckedChanged
        ' Met à jour le pourboire s'il y a lieu.

        CalculMontantTotalAPayer()

    End Sub

    Private Sub radTransport_CheckedChanged(sender As Object, e As EventArgs) Handles radTransport.CheckedChanged
        ' Met à jour le pourboire s'il y a lieu.

        CalculMontantTotalAPayer()

    End Sub

    Private Sub txtPayeP1_TextChanged(sender As Object, e As EventArgs) Handles txtPayeP1.TextChanged
        ' Met à jour le montant total payé.
        CalculMontantTotalPaye()

    End Sub

    Private Sub txtPayeP2_TextChanged(sender As Object, e As EventArgs) Handles txtPayeP2.TextChanged
        ' Met à jour le montant total payé.
        CalculMontantTotalPaye()

    End Sub

    Private Sub txtPayeP3_TextChanged(sender As Object, e As EventArgs) Handles txtPayeP3.TextChanged
        ' Met à jour le montant total payé.
        CalculMontantTotalPaye()

    End Sub

    Private Sub txtPayeP4_TextChanged(sender As Object, e As EventArgs) Handles txtPayeP4.TextChanged
        ' Met à jour le montant total payé.
        CalculMontantTotalPaye()

    End Sub

    Private Sub btnRecommencer_Click(sender As Object, e As EventArgs) Handles btnRecommencer.Click
        ' Remise à zéro du formulaire pour la dépense en cours.

        VideFormulaire()

    End Sub

    Private Sub btnRepartir_Click(sender As Object, e As EventArgs) Handles btnRepartir.Click
        ' Traite et répartie les dépenses. Garde aussi un sommaire des dépenses globales.

        ' Incrémente le compteur de dépenses de 1.
        CompteurNbDepenses()

        BloquerChangementPersonne()

        ' Ajoute le montant payé selon le type de dépense.
        Select Case DepenseType()
            Case 1
                'Restaurant
                decSommeRestaurant += decTotalPaye
            Case 2
                'Transport
                decSommeTransport += decTotalPaye
            Case 3
                'Hebergement
                decSommeHebergement += decTotalPaye
            Case 4
                'Autre
                decSommeAutre += decTotalPaye
        End Select

        decSommeDepenses = decSommeRestaurant + decSommeTransport + decSommeHebergement + decSommeAutre

        ' Si il y a plus d'une personne, répartie les dépenses.
        If NombrePersonneActive() <> 1 Then
            RepartieEtAfficheMontant()
        End If

        VideFormulaire()

    End Sub

    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ' Affiche le bilan lors de la fermeture.

        MessageBox.Show(BilanEcran(), "Bilan de la répartition", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub
End Class
