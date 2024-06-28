Imports excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.DateTime
Public Class Form1
    Dim i As Integer
    Dim q As Integer
    Dim x As Integer = 0
    Dim b As Integer = 0
    Dim p As Integer = 2
    Dim z As Integer
    Dim n, pos As Integer
    Dim numauth() As Char = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    Dim tnom(n), tanom(n) As String
    Dim tqte(n), taqte(n) As Integer
    Dim tbnom(n), tcnom(n) As String
    Dim tbper(n), tcper(n) As String
    Dim app As excel.Application
    Dim book As excel.Workbook
    Dim sheet As excel.Worksheet
    Dim sheets As excel.Worksheet
    Dim sheetes As excel.Worksheet
    Dim Shet As excel.Worksheet
    Dim range As excel.Range
    Dim l As Integer = 0
  
    Public Sub connexion_a_excel()
        Dim app As excel.Application
        Dim book As excel.Workbook
        Dim sheet As excel.Worksheet
        Dim range As excel.Range

        'Creation d'une nouvelle feuille de calcul
        app = CreateObject("excel.Application")
        app.Visible = True
        book = app.Workbooks.Add
        sheet = book.ActiveSheet


        'Ajouter des noms/valeurs dans les tableau
        sheet.Cells(1, 1).value = "NOM"
        sheet.Cells(1, 2).value = "PRENOM"
        sheet.Cells(1, 3).value = "Mots de passe"
        sheet.Cells(1, 4).value = "CONFIRMATION"

        'FORMATER UNE CELLULE EXCEL
        With sheet.Range("A1", "D1")
            .Font.Bold = True
            .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
        End With

        'CREATION D'UN TABLEAU DANS LEQUEL ON VA STOCKER PLUSIEURS VALEURS A LA FOIS
        Dim etudiant(5, 2) As String
        etudiant(0, 0) = "CHRI"
        etudiant(0, 1) = "THES"
        etudiant(1, 0) = "FRAN"
        etudiant(1, 1) = "MARI"
        etudiant(2, 0) = "ARTH"
        etudiant(2, 1) = "HARO"
        etudiant(3, 0) = "LEAN"
        etudiant(3, 1) = "KAMG"

        'REMPLIR LES INFORMATION A2:B6 DU TABLEAU DANS NOTRE FICHIERS EXCEL
        sheet.Range("A2", "B6").Value = etudiant

        'REMPLIR LES INFORMATIONS DANS C2;C6 AVEC LA FORMULE (=A2 & " " & B2).
        range = sheet.Range("C2", "C6")
        range.Formula = "=A2 & "" "" & B2"

        'REMPLIR D2:D6
        sheet.Cells(2, 4).value = "add"
        sheet.Cells(3, 4).value = "cont"
        sheet.Cells(4, 4).value = "cont"
        sheet.Cells(5, 4).value = "cont"
        sheet.Cells(6, 4).value = "cont"

        'AUTO FIT COLUMNS A:D
        range = sheet.Range("A1", "D1")
        range.EntireColumn.AutoFit()

        app.Visible = True
        app.UserControl = False

        'tout vider
        range = Nothing
        sheet = Nothing
        book = Nothing
        app.Quit()
        app = Nothing
        Exit Sub
Err_handler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub

    Public Sub ouverture_excelf1()
        app = CreateObject("excel.Application")
        app.Visible = True

        book = app.Workbooks.Open(CurDir() & "\BASE\BASE DE DONNES POUR LES PRODUITS.xlsx")
        sheet = book.Worksheets(1)

    End Sub

    Public Sub ouverture_authenti()
        app = CreateObject("excel.Application")
        app.Visible = True

        book = app.Workbooks.Open(CurDir() & "\BASE\Authentification.xlsx")
        sheet = book.Worksheets(1)
    End Sub

    Public Sub ouverture_excelf2()
        app = CreateObject("excel.Application")
        app.Visible = True

        book = app.Workbooks.Open(CurDir() & "\BASE\BASE DE DONNES POUR LES PRODUITS.xlsx")
        sheetes = book.Worksheets(2)
    End Sub

    Public Sub fermeture_excel()
        'fermer excel
        book.Save()
        range = Nothing
        sheet = Nothing
        book = Nothing
        app.Quit()
        app = Nothing
        For Each p As Process In Process.GetProcesses
            If p.ProcessName = "Microsoft Excel (32 bits)" Then
                p.Kill()
            End If
        Next
    End Sub

    Public Sub compter_les_cellules_excel()
        n = 0
        n = app.WorksheetFunction.CountIf(sheet.Columns(1), "<>")
        n = n - 1
    End Sub

    Public Sub enregistrement_produit()
        app = CreateObject("excel.Application")
        app.Visible = True

        book = app.Workbooks.Open(CurDir() & "\BASE\BASE DE DONNES POUR LES PRODUITS.xlsx")
        sheet = book.Worksheets(1)

        i = sheet.UsedRange.Rows.Count
        i = i + 1
        'ENVOI DES INFORMATIONS DANS LE TABLEUR EXCEL
        sheet.Cells(i, 1).value = TextBox3.Text
        sheet.Cells(i, 2).value = TextBox4.Text
        sheet.Cells(i, 3).value = TextBox5.Text
        sheet.Cells(i, 4).value = DateTimePicker1.Text
        sheet.Cells(i, 5).value = DateTimePicker2.Text

        sheetes = book.Worksheets(2)

        'ENVOI DES INFORMATIONS DANS LE TABLEUR EXCEL
        sheetes.Cells(i, 1).value = TextBox3.Text
        sheetes.Cells(i, 2).value = TextBox4.Text
        sheetes.Cells(i, 3).value = TextBox5.Text
        sheetes.Cells(i, 4).value = DateTimePicker1.Text
        sheetes.Cells(i, 5).value = DateTimePicker2.Text
        'tout vider

        compter_les_cellules_excel()

        Dim tab(n) As String
        For j = 2 To n + 1
            tab(x) = sheet.Cells(j, 1).value
            x = x + 1
        Next
        For m = 0 To n - 1
            ComboBox1.Items.Add(tab(m))
            ComboBox2.Items.Add(tab(m))
            ComboBox3.Items.Add(tab(m))
            ComboBox4.Items.Add(tab(m))
            ComboBox6.Items.Add(tab(m))
            ComboBox7.Items.Add(tab(m))
        Next

        'REMPLIR NOTRE COMBOBOX AVEC NOS DONNEES DE EXCEL
        book.Save()
        range = Nothing
        sheet = Nothing
        book = Nothing
        app.Quit()
        app = Nothing
        For Each p As Process In Process.GetProcesses
            If p.ProcessName = "Microsoft Excel (32 bits)" Then
                p.Kill()
            End If
        Next
    End Sub

    Public Sub achat()
        app = CreateObject("excel.Application")
        app.Visible = True

        book = app.Workbooks.Open(CurDir() & "\BASE\BASE DE DONNES POUR LES PRODUITS.xlsx")
        sheets = book.Worksheets(3)
        sheetes = book.Worksheets(2)
        sheet = book.Worksheets(1)
        'RECUPERER LES INFORMATIONS SUR LES PRODUITS DANS EXCEL ET METTTRE DANS NOTRE COMBOBOX
        q = sheets.UsedRange.Rows.Count
        q = q + 1
        'ENVOI DES INFORMATIONS DANS LE TABLEUR EXCEL
        For w = 2 To q + 1
            If ComboBox1.Text = sheets.Cells(w, 1).value Then
                sheets.Cells(w, 2).value = sheets.Cells(w, 2).value + CInt(TextBox9.Text)
            Else
                sheets.Cells(q, 1).value = ComboBox1.Text
                sheets.Cells(q, 2).value = TextBox9.Text
                sheets.Cells(q, 3).value = TextBox8.Text
                sheets.Cells(q, 4).value = DateTimePicker3.Text
            End If
        Next
        
        'COMPTER LE NOMBRE DE CELLULES EXCEL OCCUPÉ
        'compter_les_cellules_excel()

        For r = 2 To n + 1
            If ComboBox1.Text = sheetes.Cells(r, 1).value Then
                If CInt(TextBox9.Text) < sheetes.Cells(r, 2).value Then
                    sheetes.Cells(r, 2).value = sheetes.Cells(r, 2).value - CInt(TextBox9.Text)
                End If
            End If
        Next

        'FERMER EXCEL
        book.Save()
        range = Nothing
        sheet = Nothing
        book = Nothing
        app.Quit()
        app = Nothing
        For Each p As Process In Process.GetProcesses
            If p.ProcessName = "Microsoft Excel (32 bits)" Then
                p.Kill()
            End If
        Next
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim tp As Integer
        Form2.Show()
        MsgBox("BONJOUR ET BIENVENU")

        Thread.Sleep(3000)
        Form2.Close()
        GroupBox1.Parent = Me
        GroupBox2.Parent = Me
        GroupBox3.Parent = Me
        GroupBox4.Parent = Me
        GroupBox5.Parent = Me
        GroupBox6.Parent = Me
        GroupBox7.Parent = Me
        GroupBox10.Parent = Me
        GroupBox11.Parent = Me
        GroupBox10.Location = New Point(1, 1)
        GroupBox10.Size = New Point(922, 559)
        GroupBox10.BringToFront()
        Button11.Hide()
        Label62.Hide()
        Me.Size = New Point(937, 605)
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button35.Enabled = False

    End Sub

    'Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   GroupBox4.Location = New Point(1, 1)
    '  GroupBox4.Size = New Point(922, 559)
    ' GroupBox4.BringToFront()
    'End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        GroupBox6.Location = New Point(1, 1)
        GroupBox6.Size = New Point(922, 559)
        GroupBox6.BringToFront()
    End Sub

    'Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   GroupBox5.Location = New Point(1, 1)
    '  GroupBox5.Size = New Point(922, 559)
    ' GroupBox5.BringToFront()
    'End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'MONTRER LA FENETRE D'AUTHENTIFICATION

        GroupBox1.Location = New Point(1, 1)
        GroupBox1.Size = New Point(922, 559)
        GroupBox1.BringToFront()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'MONTRER LA FENETRE DE L'ENREGISTREMENT DE LA MERCHANDISES
        GroupBox2.Location = New Point(1, 1)
        GroupBox2.Size = New Point(922, 559)
        GroupBox2.BringToFront()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'MONTRER LA FENETRE DE LA SORITE DE LA MARCHANDISE
        GroupBox3.Location = New Point(1, 1)
        GroupBox3.Size = New Point(922, 559)
        GroupBox3.BringToFront()
        ouverture_excelf1()
        compter_les_cellules_excel()

        fermeture_excel()
    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'MASQUER ET AFFICHER LE MOTS DE PASSE
        TextBox2.UseSystemPasswordChar = False
        Button9.Hide()
        Button11.Show()
    End Sub

    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        'MASQUER ET AFFICHER LE MOTS DE PASSE
        TextBox2.UseSystemPasswordChar = True
        Button11.Hide()
        Button9.Show()
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim nt As Integer
        ouverture_authenti()
        nt = sheet.UsedRange.Rows.Count
        For fi = 2 To nt
            If TextBox1.Text = sheet.Cells(fi, 1).value And TextBox2.Text = sheet.Cells(fi, 2).value And sheet.Cells(fi, 5).value = "administrateur" Then
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button5.Enabled = True
                Button35.Enabled = True
                'MsgBox("VOUS ETES CONNECTÉ")
                TextBox1.Clear()
                TextBox2.Clear()
            Else
                'MsgBox("CE COMPTE N'EXISTE PAS OU VERIFIER VOTRE MOT DE PASSE")
            End If

            If TextBox1.Text = sheet.Cells(fi, 1).value And TextBox2.Text = sheet.Cells(fi, 2).value And sheet.Cells(fi, 5).value = "invite" Then
                Button2.Enabled = False
                Button3.Enabled = True
                Button4.Enabled = False
                Button5.Enabled = True
                Button39.Enabled = False
                Button41.Enabled = False
                TextBox1.Clear()
                TextBox2.Clear()
            Else
                'MsgBox("CE COMPTE N'EXISTE PAS OU VERIFIER VOTRE MOT DE PASSE")
            End If
        Next

        fermeture_excel()
        If TextBox1.Text = "" Then
            MsgBox("VOUS ETES CONNECTÉ")
        Else
            MsgBox("CE COMPTE N'EXISTE PAS OU VERIFIER VOTRE MOT DE PASSE")
        End If

        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        'VALIDES L'ENREGISTREMENT D'UNE MARCHANDISE
        If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or DateTimePicker1.Text = "" Or DateTimePicker2.Text = "" Then
            MsgBox("VEILLEZ REMPLIR TOUTES LES INFORMATIONS")
        Else
            enregistrement_produit()
        End If
        MsgBox("ENREGISTRÉ")
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        'VALIDER LA SORTIE DE OLA MARCHANDISE
        If ComboBox1.Text = "" Or TextBox9.Text = "" Or TextBox8.Text = "" Or DateTimePicker3.Text = "" Then
            MsgBox("VEILLEZ REMPLIR TOUTES LES INFORMATIONS")
        Else
            achat()
        End If
        MsgBox("ENREGISTRÉ")
        ComboBox1.Text = ""
        TextBox9.Clear()
        TextBox8.Clear()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'REMPLIR NOTRE COMBOBOX AVEC NOS DONNEES DE EXCEL
        ouverture_excelf1()
        Dim tab(n) As String
        For j = 2 To n + 1
            tab(x) = sheet.Cells(j, 1).value
            x = x + 1
        Next
        For m = 0 To n - 1
            ComboBox1.Items.Add(tab(m))
            ComboBox2.Items.Add(tab(m))
            ComboBox3.Items.Add(tab(m))
            ComboBox4.Items.Add(tab(m))
            ComboBox6.Items.Add(tab(m))
            ComboBox7.Items.Add(tab(m))
        Next
        fermeture_excel()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'ENREGISTREMENT D'UN NOUVEL UTILISATEUR
        Dim sauve, statu As String
        Dim nom = TextBox11.Text
        Dim mdp = TextBox12.Text
        Dim conf = TextBox13.Text
        Dim num = TextBox14.Text
        Dim adresse = TextBox15.Text

        If TextBox11.Text = "" And TextBox12.Text = "" And TextBox13.Text = "" And TextBox14.Text = "" And TextBox15.Text = "" Then
            Label24.Visible = True
        Else
            'VALIDER L'ENREGISTREMENT D'UN UTILISATEUR EN TANT QUE ADMIN OU INVITÉ
            Dim etat As DialogResult = MessageBox.Show("VOUS VOULEZ ENREGISTRER EN TANT QUE ADMINISTRATEUR ?", "STATUT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If etat = Windows.Forms.DialogResult.No Then
                statu = "invité"
            Else
                statu = "administrateur"
            End If
            'ENVOYER LES INFORMATIONS DANS NOTRE FICHIER TEXTE
            sauve = nom & ":" & mdp & ":" & conf & ":" & num & ":" & adresse & ":" & statu
            IO.File.WriteAllText(CurDir() & "\BASE\authentification.txt", sauve)
            MsgBox("VOS INFORMATIONS ONT ÉTÉ ENREGISTRÉ")
        End If
    End Sub

    Private Sub Button16_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        'RECUPERATION D'UN MOT DE PASSE
        For Each Ligne As String In IO.File.ReadAllLines(CurDir() & "\BASE\authentification.txt")
            If TextBox16.Text = Ligne.Split(":")(4) Then
                MsgBox(Ligne.Split(":")(2))
            Else

            End If
        Next
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        'MONTRER LA FENETRE DE LA DESTRUCTION DE LA MARCHANDISE
        GroupBox7.Location = New Point(1, 1)
        GroupBox7.Size = New Point(922, 559)
        GroupBox7.BringToFront()
    End Sub

    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        'MONTRER LA FENETRE DU CONTROLE QUALITÉ
        GroupBox9.Location = New Point(1, 1)
        GroupBox9.Size = New Point(922, 559)
        GroupBox9.BringToFront()
    End Sub

    Private Sub Button12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        'MONTRER LA FENETRE DU CONTROLE QUANTITÉ
        GroupBox8.Location = New Point(1, 1)
        GroupBox8.Size = New Point(922, 559)
        GroupBox8.BringToFront()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        'MONTRER LA FENETRE DU BILAN
        GroupBox11.Location = New Point(1, 1)
        GroupBox11.Size = New Point(922, 559)
        GroupBox11.BringToFront()
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        'AFFICHER LA PAGE DU BILAN GLOBAL
        GroupBox12.Location = New Point(1, 1)
        GroupBox12.Size = New Point(922, 559)
        GroupBox12.BringToFront()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        'VALIDATION DE LA DESTRUCTION D'UNE MARCHANDISE
        If ComboBox2.Text = "" Or TextBox6.Text = "" Or ComboBox5.Text = "" Then
            MsgBox("VEUILLEZ REMPLIR TOUTES LES INFORMATIONS NECCESAIRE")
        Else
            ouverture_excelf2()
            Shet = book.Worksheets(4)

            z = Shet.UsedRange.Rows.Count
            z = z + 1

            Shet.Cells(z, 1) = ComboBox2.Text
            Shet.Cells(z, 2) = TextBox6.Text
            Shet.Cells(z, 3) = ComboBox5.Text

            For g = 2 To n + 1
                If ComboBox2.Text = sheetes.Cells(g, 1).value Then
                    If sheetes.Cells(g, 2).value > CInt(TextBox6.Text) Then
                        sheetes.Cells(g, 2).value = sheetes.Cells(g, 2).value - CInt(TextBox6.Text)
                        'MsgBox("LA QUANTITÉ SPECIFIÉ DE LA MARCHANDISE A ÉTÉ DÉTRUITE")
                    Else
                        MsgBox("LA QUANTITÉ QUE VOUS VOULEZ SUPPRIMER EST INFÉRIEUR AU STOCK DISPONIBLE")
                        MsgBox("PENSEZ A REFAIRE VOTRE STOCK")
                    End If
                End If
            Next
        End If
        fermeture_excel()
        TextBox6.Clear()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        'CONTROLE QUALITÉ D'UN PRODUIT
        Dim f As Integer = 0
        If ComboBox4.Text = "" Then
            MsgBox("VEUILLES REMPLIR TOUTES LES INFORMATIONS NECCESAIRES")
        Else
            ouverture_excelf2()
            For g = 2 To n + 1
                If ComboBox4.Text = sheetes.Cells(g, 1).value Then
                    tbnom(f) = sheetes.Cells(g, 1).value
                    tbper(f) = sheetes.Cells(g, 5).value
                End If
            Next
            f = f + 1
            fermeture_excel()
        End If
        DataGridView2.Rows.Clear()
        For g = 0 To f - 1
            DataGridView2.Rows.Add({1, tbnom(g), tbper(g)})
        Next
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        'CONTROLE QUALITÉ DE TOUT LE STOCK
        ouverture_excelf2()
        DataGridView2.Rows.Clear()
        For u = 2 To n + 1
            DataGridView2.Rows.Add({u - 1, sheetes.Cells(u, 1).value, sheetes.Cells(u, 5).value})
        Next
        fermeture_excel()
    End Sub

    Private Sub Button13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        'CONTROLE QUANTITÉ D'UNE SEUL MARCHANDISE
        Dim b As Integer = 0
        If ComboBox3.Text = "" Then
            MsgBox("VEUILLEZ REMPLIR TOUTES LES INFORMATIONS NECCESAIRES")
        Else
            ouverture_excelf2()
            For g = 2 To n + 1
                If ComboBox3.Text = sheetes.Cells(g, 1).value Then
                    tnom(b) = sheetes.Cells(g, 1).value
                    tqte(b) = sheetes.Cells(g, 2).value
                End If
            Next
            b = b + 1
            fermeture_excel()
        End If
        DataGridView1.Rows.Clear()
        For g = 0 To b - 1
            DataGridView1.Rows.Add({1, tnom(g), tqte(g)})
        Next
    End Sub

    Private Sub Button15_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        'CONTROLE QUANTITÉ DE TOUT LE STOCK
        ouverture_excelf2()
        DataGridView1.Rows.Clear()
        For u = 2 To n + 1
            DataGridView1.Rows.Add({u - 1, sheetes.Cells(u, 1).value, sheetes.Cells(u, 2).value})
        Next
        fermeture_excel()
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'VALIDATION DE LA FERMETURE DE L'APPLICATION
        Dim fermer As DialogResult = MessageBox.Show("VOULEZ VOUS VRAIMENT QUITTER ?", "FERMER", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If fermer = Windows.Forms.DialogResult.Yes Then
            Me.Close()
        Else

        End If

    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        'RENTRER AU FORME DE BIENVENU
        GroupBox10.BringToFront()
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click, Button25.Click
        'RENTRER AU FORME DE BIENVENU
        GroupBox10.BringToFront()
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click, Button26.Click
        'RENTRER AU FORME DE BIENVENU
        GroupBox10.BringToFront()
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click, Button27.Click
        'RENTRER AU FORME DE BIENVENU
        GroupBox10.BringToFront()
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        'RENTRER AU FORME D'AUTHENTIFICATION
        GroupBox1.BringToFront()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click, Button29.Click
        'RENTRER AU FORME D'AUTHENTIFICATION
        GroupBox1.BringToFront()
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click, Button30.Click
        'RETOUR A LA FENETRE DE CONTROLE ET MISE A JOUR DU STOCK
        GroupBox6.BringToFront()
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click, Button31.Click
        'RETOUR A LA FENETRE DE CONTROLE ET MISE A JOUR DU STOCK
        GroupBox6.BringToFront()
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click, Button32.Click
        'RETOUR A LA FENETRE DE CONTROLE ET MISE A JOUR DU STOCK
        GroupBox6.BringToFront()
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        'RETOUR A LA FENETRE DE CONTROLE ET MISE A JOUR DU STOCK
        GroupBox10.BringToFront()
    End Sub

    Private Sub Form1_SizeChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
        'AJUSTEMENT AUTOMATIQUE DE LA TAILLE
        GroupBox1.Size = Me.Size
        GroupBox2.Size = Me.Size
        GroupBox3.Size = Me.Size
        GroupBox4.Size = Me.Size
        GroupBox5.Size = Me.Size
        GroupBox6.Size = Me.Size
        GroupBox7.Size = Me.Size
        GroupBox8.Size = Me.Size
        GroupBox9.Size = Me.Size
        GroupBox10.Size = Me.Size
        GroupBox11.Size = Me.Size
        GroupBox12.Size = Me.Size
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress, TextBox5.KeyPress
        If Not numauth.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not numauth.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox8_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not numauth.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox9_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not numauth.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox6_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not numauth.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Public Structure bilan
        Dim nom As String
        Dim qe As Integer
        Dim qs As Integer
        Dim de As Date
        Dim ds As Date
        Dim qp As Integer
        Dim qr As Integer
    End Structure

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        Dim dd As Date
        Dim df As Date
        Dim datedeb As Date
        Dim datefin As Date
        Dim resultat1, resultats2 As Integer
        Dim con, cont, conte As Integer
        Dim co As Integer = 0
        Dim s As Integer = 2
        Dim t(7) As bilan
        dd = DateTimePicker4.Text
        df = DateTimePicker5.Text

        ouverture_excelf1()
        sheetes = book.Worksheets(2)
        sheets = book.Worksheets(3)
        Shet = book.Worksheets(4)
        n = sheet.UsedRange.Rows.Count
        n = n - 1
        con = sheets.UsedRange.Rows.Count
        con = con - 1
        cont = sheetes.UsedRange.Rows.Count
        cont = cont - 1
        conte = Shet.UsedRange.Rows.Count
        conte = conte - 1

        For s = 2 To n + 1
            datedeb = CDate(sheet.Cells(s, 4).value)
            datefin = CDate(sheet.Cells(s, 4).value)
            resultat1 = DateTime.Compare(dd, datedeb)
            resultats2 = DateTime.Compare(df, datefin)
            If resultat1 < 0 And resultats2 > 0 Then
                t(co).nom = sheet.Cells(s, 1).value
                sheet = book.Worksheets(1)
                For sa = 2 To n + 1
                    If t(co).nom = sheet.Cells(sa, 1).value Then
                        t(co).qe = sheet.Cells(sa, 2).value
                        t(co).de = CDate(sheet.Cells(sa, 4).value)
                    End If
                Next
            End If
            If resultat1 < 0 And resultats2 > 0 Then
                sheets = book.Worksheets(3)
                For sas = 2 To con + 1
                    If t(co).nom = sheets.Cells(sas, 1).value Then
                        t(co).qs = sheets.Cells(sas, 2).value
                        t(co).ds = CDate(sheets.Cells(sas, 4).value)
                    End If
                Next
            End If
            If resultat1 < 0 And resultats2 > 0 Then
                Shet = book.Worksheets(4)
                For bi = 2 To conte + 1
                    If t(co).nom = Shet.Cells(bi, 1).value Then
                        t(co).qp = Shet.Cells(bi, 2).value
                    End If
                Next
            End If
            If resultat1 < 0 And resultats2 > 0 Then
                sheetes = book.Worksheets(2)
                For biy = 2 To cont + 1
                    If t(co).nom = sheetes.Cells(biy, 1).value Then
                        t(co).qr = sheetes.Cells(biy, 2).value
                    End If
                Next
            End If
            co = co + 1
        Next
        fermeture_excel()

        DataGridView3.Rows.Clear()
        For biya = 0 To co - 1
            DataGridView3.Rows.Add({t(biya).nom, t(biya).qe, t(biya).qs, t(biya).de, t(biya).ds, t(biya).qp, t(biya).qr})
        Next
    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        'RETOUR A LA PAGE DU BILAN
        GroupBox11.BringToFront()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        GroupBox13.Location = New Point(1, 1)
        GroupBox13.Size = New Point(922, 559)
        GroupBox13.BringToFront()
        TabControl1.Hide()
        TabControl2.Hide()
    End Sub

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button47.Click
        GroupBox10.BringToFront()
    End Sub

    Public Structure remplace
        Dim pro As String
        Dim qua As Integer
        Dim pra As Integer
        Dim daa As Date
        Dim dap As Date
    End Structure

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        Dim m As Integer = 0
        Dim ni As Integer
        ouverture_excelf1()
        n = sheet.UsedRange.Rows.Count
        n = n - 1
        Dim ma(n) As remplace
        For u = 2 To n + 1
            If ComboBox6.Text = sheet.Cells(u, 1).value Then
                sheet.Cells(u, 1).value.clear()
                sheet.Cells(u, 2).value.clear()
                sheet.Cells(u, 3).value.clear()
                sheet.Cells(u, 4).value.clear()
                sheet.Cells(u, 5).value.clear()
                ni = u
            Else
                MsgBox("Veuillez essayer encore")
            End If
        Next
        ni = ni + 1
        For pa = ni To n - (ni - 1)
            ma(m).pro = sheet.Cells(ni, 1).value
            ma(m).qua = sheet.Cells(ni, 2).value
            ma(m).pra = sheet.Cells(ni, 3).value
            ma(m).daa = sheet.Cells(ni, 4).value
            ma(m).dap = sheet.Cells(ni, 5).value
            m = m + 1
        Next
        ni = ni - 1
        For pat = ni To n - ni
            sheet.Cells(ni, 1) = ma(m).pro
            sheet.Cells(ni, 2) = ma(m).qua
            sheet.Cells(ni, 3) = ma(m).pra
            sheet.Cells(ni, 4) = ma(m).daa
            sheet.Cells(ni, 5) = ma(m).dap
            m = m + 1
        Next
        fermeture_excel()
    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        ouverture_excelf1()
        sheets = book.Worksheets(2)
        n = sheet.UsedRange.Rows.Count
        For ip = 2 To n + 1
            If ComboBox7.Text = sheet.Cells(ip, 1).value Then
                sheet.Cells(ip, 1) = ComboBox7.Text
                sheet.Cells(ip, 2) = TextBox31.Text
                sheet.Cells(ip, 3) = TextBox32.Text
                sheet.Cells(ip, 4) = TextBox33.Text
                sheet.Cells(ip, 5) = TextBox34.Text
            End If
        Next
        fermeture_excel()
        MsgBox("MODIFICATION ÉFFECTUÉ")
    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        TabControl1.Location = New Point(367, 29)
        TabControl1.Size = New Point(436, 440)
        TabControl1.Show()
        TabControl1.BringToFront()
        Label34.Hide()
        Button42.Enabled = False
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        TabControl2.Location = New Point(367, 29)
        TabControl2.Size = New Point(436, 440)
        TabControl2.Show()
        TabControl2.BringToFront()
    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        Dim nat As Integer


        If TextBox19.Text = "" Or TextBox18.Text = "" Or TextBox17.Text = "" Or TextBox10.Text = "" Or TextBox7.Text = "" Then
            Label34.Show()
        ElseIf TextBox18.Text <> TextBox17.Text Then
            MsgBox("LES MOTS DE PASSE NE CORRESPONDENT PAS")
        Else
            app = CreateObject("excel.Application")
            app.Visible = True

            book = app.Workbooks.Open(CurDir() & "\BASE\Authentification.xlsx")
            sheet = book.Worksheets(1)
            nat = sheet.UsedRange.Rows.Count
            nat = nat + 1
            sheet.Cells(nat, 1) = TextBox19.Text
            sheet.Cells(nat, 2) = TextBox18.Text
            sheet.Cells(nat, 3) = TextBox10.Text
            sheet.Cells(nat, 4) = TextBox7.Text
            nat = nat + 1
            fermeture_excel()
        End If

    End Sub

    Private Sub Button1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.MouseHover
        Label62.Show()
        Label62.Text = "AUTHENTIFICATION"
    End Sub
    Private Sub Button1_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Button4.MouseLeave, Button35.MouseLeave, Button5.MouseLeave, Button6.MouseLeave, Button7.MouseLeave
        Label62.Hide()
    End Sub

    Private Sub Button2_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.MouseHover
        Label62.Show()
        Label62.Text = "ENREGISTREMENT"
    End Sub

    Private Sub Button3_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.MouseHover
        Label62.Show()
        Label62.Text = "CONTROLE STOCK"
    End Sub

    Private Sub Button4_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.MouseHover
        Label62.Show()
        Label62.Text = "SORTI DE STOCK"
    End Sub

    Private Sub Button35_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.MouseHover
        Label62.Show()
        Label62.Text = "BILAN"
    End Sub

    Private Sub Button5_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.MouseHover
        Label62.Show()
        Label62.Text = "PARAMETRES"
    End Sub

    Private Sub Button6_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.MouseHover
        Label62.Show()
        Label62.Text = "INFORMATION"
    End Sub

    Private Sub Button7_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.MouseHover
        Label62.Show()
        Label62.Text = "QUITTER"
    End Sub

    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button48.Click

        Dim at As Integer
        If TextBox21.Text = "" Or TextBox20.Text = "" Then
            MsgBox("VEUILLEZ REMPLIR TOUTES LES INFORMATIONS")
        Else
            ouverture_authenti()
            at = sheet.Cells.usedrange
            For ra = 2 To at + 1
                If TextBox21.Text = sheet.Cells(ra, 1).value And TextBox20.Text = sheet.Cells(ra, 2).value Then
                    Button42.Enabled = True
                    pos = ra
                End If
            Next
        End If
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        'MODIFICATIONS DES INFORMATIONS D'UN COMPTE
        ouverture_authenti()
        sheet.Cells(pos, 1) = TextBox26.Text
        sheet.Cells(pos, 2) = TextBox25.Text
        sheet.Cells(pos, 3) = TextBox24.Text
        sheet.Cells(pos, 4) = TextBox23.Text
        sheet.Cells(pos, 5) = TextBox22.Text
        fermeture_excel()
        MsgBox("LES INFORMATIONS ONT ÉTÉ MODIFIÉ")
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        Dim ath As Integer
        ouverture_authenti()
        ath = sheet.UsedRange.Rows.Count
        For ip = 2 To ath + 1
            If TextBox28.Text = sheet.Cells(ip, 1).value And TextBox27.Text = sheet.Cells(ip, 2).value Then
                sheet.Rows(ip).clear()
            Else
                MsgBox("VERIFIER VOTRE MOTS DE PASSE OU NOM D'UTILISATEUR")
            End If
        Next
        fermeture_excel()
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        Dim mo As Integer
        For it = 2 To mo + 1
            If TextBox29.Text = sheet.Cells(it, 1).value And TextBox30.Text = sheet.Cells(it, 4).value Then
                MsgBox("VOTRE MOT DE PASSE EST" & sheet.Cells(it, 2).value)
            Else
                MsgBox("VEUILLEZ VERIFIER CE QUE VOUS AVEZ ENTRÉ")
            End If
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Process.Start(CurDir() & "\BASE\INFORMATIONS.docx")
    End Sub

    Private Sub Button22_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        GroupBox1.BringToFront()
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button35.Enabled = False
    End Sub
End Class
