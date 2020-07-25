Option Compare Text

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnAbrir As System.Windows.Forms.Button
    Friend WithEvents cmbPara As System.Windows.Forms.ComboBox
    Friend WithEvents btnConverte As System.Windows.Forms.Button
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lvAvisos As System.Windows.Forms.ListView
    Friend WithEvents lvArquivos As System.Windows.Forms.ListView
    Friend WithEvents Imgs As System.Windows.Forms.ImageList
    Friend WithEvents btnLimpar As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgAviso As System.Windows.Forms.ImageList
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents llMail As System.Windows.Forms.LinkLabel
    Friend WithEvents lblAguarde As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnAbrir = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmbPara = New System.Windows.Forms.ComboBox
        Me.btnConverte = New System.Windows.Forms.Button
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.lvAvisos = New System.Windows.Forms.ListView
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.imgAviso = New System.Windows.Forms.ImageList(Me.components)
        Me.lvArquivos = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Imgs = New System.Windows.Forms.ImageList(Me.components)
        Me.btnLimpar = New System.Windows.Forms.Button
        Me.llMail = New System.Windows.Forms.LinkLabel
        Me.lblAguarde = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Arquivos:"
        '
        'btnAbrir
        '
        Me.btnAbrir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAbrir.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAbrir.Image = CType(resources.GetObject("btnAbrir.Image"), System.Drawing.Image)
        Me.btnAbrir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAbrir.Location = New System.Drawing.Point(465, 25)
        Me.btnAbrir.Name = "btnAbrir"
        Me.btnAbrir.Size = New System.Drawing.Size(120, 38)
        Me.btnAbrir.TabIndex = 0
        Me.btnAbrir.Text = "&Adicionar..."
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(465, 238)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Converter para:"
        '
        'cmbPara
        '
        Me.cmbPara.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbPara.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPara.Location = New System.Drawing.Point(465, 253)
        Me.cmbPara.Name = "cmbPara"
        Me.cmbPara.Size = New System.Drawing.Size(123, 21)
        Me.cmbPara.TabIndex = 2
        '
        'btnConverte
        '
        Me.btnConverte.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnConverte.Enabled = False
        Me.btnConverte.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnConverte.Image = CType(resources.GetObject("btnConverte.Image"), System.Drawing.Image)
        Me.btnConverte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnConverte.Location = New System.Drawing.Point(465, 281)
        Me.btnConverte.Name = "btnConverte"
        Me.btnConverte.Size = New System.Drawing.Size(120, 38)
        Me.btnConverte.TabIndex = 3
        Me.btnConverte.Text = "Converter!"
        '
        'OFD
        '
        Me.OFD.Filter = "Bitmaps (*.bmp *.jpg *.jpeg *.gif *.png *.ico)|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.i" & _
        "co|Qualquer aquivo|*.*"
        Me.OFD.Multiselect = True
        Me.OFD.Title = "Selecionar Bitmaps..."
        '
        'lvAvisos
        '
        Me.lvAvisos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvAvisos.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lvAvisos.LargeImageList = Me.imgAviso
        Me.lvAvisos.Location = New System.Drawing.Point(12, 324)
        Me.lvAvisos.Name = "lvAvisos"
        Me.lvAvisos.Size = New System.Drawing.Size(570, 87)
        Me.lvAvisos.SmallImageList = Me.imgAviso
        Me.lvAvisos.StateImageList = Me.imgAviso
        Me.lvAvisos.TabIndex = 5
        Me.lvAvisos.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Tipo"
        Me.ColumnHeader3.Width = 111
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Mensagem"
        Me.ColumnHeader4.Width = 434
        '
        'imgAviso
        '
        Me.imgAviso.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgAviso.ImageStream = CType(resources.GetObject("imgAviso.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgAviso.TransparentColor = System.Drawing.Color.Transparent
        '
        'lvArquivos
        '
        Me.lvArquivos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvArquivos.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvArquivos.LargeImageList = Me.Imgs
        Me.lvArquivos.Location = New System.Drawing.Point(12, 21)
        Me.lvArquivos.Name = "lvArquivos"
        Me.lvArquivos.Size = New System.Drawing.Size(444, 297)
        Me.lvArquivos.SmallImageList = Me.Imgs
        Me.lvArquivos.StateImageList = Me.Imgs
        Me.lvArquivos.TabIndex = 4
        Me.lvArquivos.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Arquivo"
        Me.ColumnHeader1.Width = 150
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Caminho"
        Me.ColumnHeader2.Width = 503
        '
        'Imgs
        '
        Me.Imgs.ImageSize = New System.Drawing.Size(16, 16)
        Me.Imgs.TransparentColor = System.Drawing.Color.Transparent
        '
        'btnLimpar
        '
        Me.btnLimpar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLimpar.Enabled = False
        Me.btnLimpar.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnLimpar.Image = CType(resources.GetObject("btnLimpar.Image"), System.Drawing.Image)
        Me.btnLimpar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnLimpar.Location = New System.Drawing.Point(465, 66)
        Me.btnLimpar.Name = "btnLimpar"
        Me.btnLimpar.Size = New System.Drawing.Size(120, 38)
        Me.btnLimpar.TabIndex = 1
        Me.btnLimpar.Text = "&Limpar lista"
        '
        'llMail
        '
        Me.llMail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llMail.Location = New System.Drawing.Point(468, 159)
        Me.llMail.Name = "llMail"
        Me.llMail.Size = New System.Drawing.Size(114, 30)
        Me.llMail.TabIndex = 8
        Me.llMail.TabStop = True
        Me.llMail.Text = "Marcius C. Bezerra (Opine!!)"
        Me.llMail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAguarde
        '
        Me.lblAguarde.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAguarde.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAguarde.ForeColor = System.Drawing.Color.Red
        Me.lblAguarde.Location = New System.Drawing.Point(468, 210)
        Me.lblAguarde.Name = "lblAguarde"
        Me.lblAguarde.Size = New System.Drawing.Size(114, 15)
        Me.lblAguarde.TabIndex = 9
        Me.lblAguarde.Text = "Aguarde..."
        Me.lblAguarde.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblAguarde.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(594, 421)
        Me.Controls.Add(Me.lblAguarde)
        Me.Controls.Add(Me.llMail)
        Me.Controls.Add(Me.btnLimpar)
        Me.Controls.Add(Me.lvArquivos)
        Me.Controls.Add(Me.lvAvisos)
        Me.Controls.Add(Me.btnConverte)
        Me.Controls.Add(Me.cmbPara)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnAbrir)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "MCB ConvGraf - Conversor Gráfico"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private WithEvents Conv As Conversao

    Private Sub btnConverte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConverte.Click
        Dim Thrd As New System.Threading.Thread(AddressOf ConverteArquivos)
        Thrd.Start()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbPara.Items.Add("GIF")
        cmbPara.Items.Add("JPEG")
        cmbPara.Items.Add("WMF")
        cmbPara.Items.Add("BMP")
        cmbPara.Items.Add("ICO")
        cmbPara.Items.Add("EMF")
        cmbPara.Items.Add("PNG")
        cmbPara.SelectedIndex = 1
        llMail.Links.Clear()
        llMail.Links.Add(0, 18, "mailto:marciusbezerra@aol.com")
    End Sub

    Private Sub btnAbrir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbrir.Click
        Try
            If OFD.ShowDialog = DialogResult.OK Then
                Dim Arq As String
                For Each Arq In OFD.FileNames
                    lvArquivos.Items.Add(System.IO.Path.GetFileName(Arq), _
                        Imgs.Images.Add(New System.Drawing.Bitmap(Arq), _
                        Imgs.TransparentColor)).SubItems.Add(Arq)
                Next
            End If
        Finally
            HabilitaBotoes(False)
        End Try
    End Sub

    Public Sub ConverteArquivos()
        If lvArquivos.Items.Count = 0 Then
            MsgBox("É preciso selecionar arquivos primeiro, clicando em Adicionar...", _
                MsgBoxStyle.Information, Text)
            btnAbrir.Focus()
            Exit Sub
        End If
        Dim Arq As ListViewItem
        Dim Formato As Conversao.ParaTipo
        Select Case cmbPara.SelectedIndex
            Case 0 : Formato = Conversao.ParaTipo.GIF
            Case 1 : Formato = Conversao.ParaTipo.JPEG
            Case 2 : Formato = Conversao.ParaTipo.WMF
            Case 3 : Formato = Conversao.ParaTipo.BMP
            Case 4 : Formato = Conversao.ParaTipo.ICO
            Case 5 : Formato = Conversao.ParaTipo.EMF
            Case 6 : Formato = Conversao.ParaTipo.PNG
        End Select
        Try
            HabilitaBotoes(True)
            lblAguarde.Show()
            lvAvisos.Items.Clear()
            Application.DoEvents()
            Conv = New Conversao
            For Each Arq In lvArquivos.Items
                Conv.Converte(Arq.SubItems(0).Text, Formato)
            Next
            MsgBox("Conversão de arquivo(s) finalizada.", _
                MsgBoxStyle.Information, Text)
        Catch ex As Exception
            MsgBox("Erro: " & ex.Message, MsgBoxStyle.Critical, Text)
        Finally
            lblAguarde.Hide()
            HabilitaBotoes(False)
            If Not Conv Is Nothing Then Conv = Nothing
        End Try
    End Sub

    Private Sub btnLimpar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpar.Click
        lvArquivos.Items.Clear()
        Imgs.Images.Clear()
        HabilitaBotoes(False)
    End Sub

    Private Sub Conv_Status(ByVal Tipo As String, ByVal Msg As String, ByVal Imagem As Integer) Handles Conv.Status
        Dim it As ListViewItem
        it = lvAvisos.Items.Add(Tipo, Imagem)
        it.SubItems.Add(Msg)
        it.EnsureVisible()
    End Sub

    Private Sub HabilitaBotoes(ByVal Processando As Boolean)
        btnAbrir.Enabled = Not Processando
        cmbPara.Enabled = Not Processando
        btnConverte.Enabled = ((lvArquivos.Items.Count > 0) And Not Processando)
        btnLimpar.Enabled = ((lvArquivos.Items.Count > 0) And Not Processando)
    End Sub

    Private Sub llMail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llMail.LinkClicked
        System.Diagnostics.Process.Start(CType(e.Link.LinkData, String))
        e.Link.Visited = True
    End Sub
End Class
