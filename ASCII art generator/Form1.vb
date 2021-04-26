Imports System.IO
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ofd As New OpenFileDialog
        ofd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
        ofd.Filter = "JPEG files (*.jpg)|*.jpg|Bitmap files (*.bmp)|*.bmp"
        Dim result As DialogResult = ofd.ShowDialog
        If Not (PictureBox1) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox1.BackgroundImage = Image.FromFile(ofd.FileName)
            PictureBox1.BackgroundImageLayout = ImageLayout.Zoom
        End If
        ImageFunctions()
    End Sub

    Sub ImageFunctions()
        Dim bmp As Bitmap = PictureBox1.BackgroundImage
        PictureBox2.BackgroundImage = ConvertToGreyscale(bmp)
        PictureBox2.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Function ConvertToGreyscale(ByVal source As Bitmap) As Bitmap
        Dim bm As New Bitmap(source)
        Using fp As New FastPix(bm)
            For y As Integer = 0 To bm.Height - 1
                For x As Integer = 0 To bm.Width - 1
                    Dim c As Color = fp.GetPixel(x, y)
                    Dim luma As Integer = CInt(c.R * 0.3 + c.G * 0.59 + c.B * 0.11)
                    fp.SetPixel(x, y, Color.FromArgb(luma, luma, luma))
                Next
            Next
        End Using
        Return bm
    End Function


End Class
