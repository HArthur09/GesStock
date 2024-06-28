Public Class stock
    Public nom_du_produit As String
    Public quantité As Integer
    Public prix_acht As Integer
    Public date_achat As String
    Public peremption As String

    Public Sub New()
        nom_du_produit = ""
        quantité = 0
        prix_acht = 0
        date_achat = ""
        peremption = ""
    End Sub
End Class
