Public Class Principal

    Private Sub ClientesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClientesToolStripMenuItem.Click
        Cliente.MdiParent = Me
        Cliente.Show()
    End Sub

    Private Sub MovimentoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MovimentoToolStripMenuItem.Click
        Movimento.MdiParent = Me
        Movimento.Show()
    End Sub

    Private Sub Devolu��oToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Devolu��oToolStripMenuItem.Click
        My.Forms.Devolucao.MdiParent = Me
        My.Forms.Devolucao.Show()
    End Sub

    Private Sub IncluirCategoriasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IncluirCategoriasToolStripMenuItem.Click
        Categorias.MdiParent = Me
        Categorias.Show()
    End Sub

    Private Sub TodasAsCategoriasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TodasAsCategoriasToolStripMenuItem.Click
        TodasCategorias.MdiParent = Me
        TodasCategorias.Show()
    End Sub
End Class
