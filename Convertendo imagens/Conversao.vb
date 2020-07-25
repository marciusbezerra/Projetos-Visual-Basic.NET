

Public Class Conversao

    Public Enum ParaTipo
        JPEG
        GIF
        BMP
        ICO
        WMF
        PNG
        EMF
    End Enum

    Public Event Status(ByVal Tipo As String, ByVal Msg As String, ByVal Imagem As Integer)

    Public Sub Converte(ByVal Arquivo As String, ByVal ParaOTipo As ParaTipo)
        If Not System.IO.File.Exists(Arquivo) Then
            RaiseEvent Status("Erro", "Arquivo '" & Arquivo & "' não localizado.", 0)
            Exit Sub
        End If
        Dim b As System.Drawing.Bitmap
        Try
            b = New System.Drawing.Bitmap(Arquivo)
            Dim ArquivoDestino = System.IO.Path.GetFileNameWithoutExtension(Arquivo)
            Dim Converteu As Boolean = False
            Select Case ParaOTipo
                Case ParaTipo.GIF
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".gif", _
                        System.Drawing.Imaging.ImageFormat.Gif)
                Case ParaTipo.JPEG
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".jpg", _
                        System.Drawing.Imaging.ImageFormat.Jpeg)
                Case ParaTipo.WMF
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".wmf", _
                        System.Drawing.Imaging.ImageFormat.Wmf)
                Case ParaTipo.BMP
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".bmp", _
                        System.Drawing.Imaging.ImageFormat.Bmp)
                Case ParaTipo.ICO
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".ico", _
                        System.Drawing.Imaging.ImageFormat.Icon)
                Case ParaTipo.PNG
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".png", _
                        System.Drawing.Imaging.ImageFormat.Png)
                Case ParaTipo.EMF
                    Converteu = ConverSeNaoExite(b, ArquivoDestino & ".emf", _
                        System.Drawing.Imaging.ImageFormat.Emf)
            End Select
            If Converteu Then _
                RaiseEvent Status("Sucesso", _
                    "Arquivo '" & Arquivo & "' convertido com sucesso.", 1)
        Catch ex As ArgumentException
            RaiseEvent Status("Erro", "'" & Arquivo & "' parece não ser um bitmap válido.", 0)
        Catch ex As Exception
            RaiseEvent Status("Erro", ex.Message, 0)
        Finally
            If Not b Is Nothing Then b.Dispose()
        End Try
    End Sub

    Private Function ConverSeNaoExite(ByVal b As System.Drawing.Bitmap, _
        ByVal Arquivo As String, _
        ByVal Formato As System.Drawing.Imaging.ImageFormat) As Boolean
        Dim RetBool As Boolean = False
        If System.IO.File.Exists(Arquivo) Then
            RaiseEvent Status("Erro", "'" & Arquivo & "' já existe.", 0)
        Else
            b.Save(Arquivo, Formato)
            RetBool = True
        End If
        Return RetBool
    End Function
End Class
