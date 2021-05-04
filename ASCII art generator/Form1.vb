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

        bmp = PictureBox2.BackgroundImage
        PictureBox2.BackgroundImage = CropToScale(bmp)
        PictureBox2.BackgroundImageLayout = ImageLayout.Zoom

        'bmp = PictureBox2.BackgroundImage
        'PictureBox2.BackgroundImage = Pixelate(bmp)
        'PictureBox2.BackgroundImageLayout = ImageLayout.Zoom
    End Sub



    Function CropToScale(ByVal source As Bitmap) As Bitmap
        Dim CropRect As New Rectangle(100, 0, 100, 100)
        Dim OriginalImage = source
        Dim CropImage = New Bitmap(CropRect.Width, CropRect.Height)
        Using grp = Graphics.FromImage(CropImage)
            grp.DrawImage(OriginalImage, New Rectangle(0, 0, CropRect.Width, CropRect.Height), CropRect, GraphicsUnit.Pixel)
            OriginalImage.Dispose()
        End Using
        Return CropImage
    End Function


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


    Function Pixelate(ByVal source As Bitmap) As Bitmap
        Dim bm As New Bitmap(source)
        Dim c As Color
        Dim averageRed As Integer
        Dim averageGreen As Integer
        Dim averageBlue As Integer
        Dim x, y As Integer

        Using fp As New FastPix(bm)


            For yCount = 0 To Math.Floor(source.Height / 50)
                For xCount = 0 To Math.Floor(source.Width / 50)


                    For y = 0 To 49
                        For x = 0 To 49
                            c = fp.GetPixel(x + xCount * 50, y + yCount * 50)
                            averageRed += c.R
                            averageGreen += c.G
                            averageBlue += c.B
                        Next
                    Next

                    averageRed /= (x * y)
                    averageGreen /= (x * y)
                    averageBlue /= (x * y)

                    For y = 0 To 49
                        For x = 0 To 49
                            fp.SetPixel(x + xCount * 50, y + yCount * 50, Color.FromArgb(averageRed, averageGreen, averageBlue))
                        Next
                    Next
                    averageRed = 0
                    averageGreen = 0
                    averageBlue = 0
                Next
            Next


            Return bm


        End Using
    End Function

End Class
