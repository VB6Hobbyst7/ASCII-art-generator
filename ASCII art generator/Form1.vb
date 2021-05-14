Imports System.IO
Imports System.Threading

Public Class Form1


    Public pixelSize As Integer
    Public Value As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ofd As New OpenFileDialog
        ofd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
        ofd.Filter = "JPEG files (*.jpg)|*.jpg|Bitmap files (*.bmp)|*.bmp"
        Dim result As DialogResult = ofd.ShowDialog
        If Not (PictureBox1) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox1.BackgroundImage = Image.FromFile(ofd.FileName)
            PictureBox1.BackgroundImageLayout = ImageLayout.Zoom
        End If
        pixelSize = Int(TrackBar1.Value)
        ImageFunctions()
    End Sub

    Sub ImageFunctions()
        Dim asciiart(,) As String


        Dim bmp As Bitmap = PictureBox1.BackgroundImage
        PictureBox2.BackgroundImage = ConvertToGreyscale(bmp)
        PictureBox2.BackgroundImageLayout = ImageLayout.Zoom

        bmp = PictureBox2.BackgroundImage
        PictureBox2.BackgroundImage = CropToScale(bmp)
        PictureBox2.BackgroundImageLayout = ImageLayout.Zoom

        TopThread(bmp)

        bmp = PictureBox2.BackgroundImage
        PictureBox2.BackgroundImage = Pixelate(bmp)
        PictureBox2.BackgroundImageLayout = ImageLayout.Zoom

        bmp = PictureBox2.BackgroundImage
        asciiart = ConvertToText(bmp)

        Dim finalArt As String


        For y = 0 To (bmp.Height - 1) / pixelSize
            For x = 0 To (bmp.Width - 1) / pixelSize
                finalArt = finalArt + asciiart(x, y)
            Next
            finalArt += vbNewLine
        Next

        TextBox1.Text = finalArt
    End Sub



    Function CropToScale(ByVal source As Bitmap) As Bitmap
        Dim bm As New Bitmap(source)
        Dim newWidth As Integer = bm.Width - (bm.Width Mod pixelSize)
        Dim newHeight As Integer = bm.Height - (bm.Height Mod pixelSize)

        Dim CropRect As New Rectangle()
        CropRect.Width = newWidth
        CropRect.Height = newHeight
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

            For yCount = 0 To (source.Height / pixelSize) - 1
                For xCount = 0 To (source.Width / pixelSize) - 1

                    For y = 0 To pixelSize - 1
                        For x = 0 To pixelSize - 1
                            c = fp.GetPixel(x + xCount * pixelSize, y + yCount * pixelSize)
                            averageRed += c.R
                            averageGreen += c.G
                            averageBlue += c.B
                        Next
                    Next

                    averageRed /= (x * y)
                    averageGreen /= (x * y)
                    averageBlue /= (x * y)

                    For y = 0 To pixelSize - 1
                        For x = 0 To pixelSize - 1
                            fp.SetPixel(x + xCount * pixelSize, y + yCount * pixelSize, Color.FromArgb(averageRed, averageGreen, averageBlue))
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

    Function TopThread(ByVal source As Bitmap) As Bitmap

        Dim newThread As New Counter(Me, 0)

        Dim threadOne, threadTwo As New Thread(AddressOf newThread.Run)

        threadOne.IsBackground = True
        threadTwo.IsBackground = True
        threadOne.Start()
        threadTwo.Start()




    End Function


    Function ConvertToText(ByVal source As Bitmap) As String(,)
        Dim bm As New Bitmap(source)
        Dim brightness As Double
        Dim position As Integer
        Dim asciiArt(bm.Width - (bm.Width Mod pixelSize), bm.Height - (bm.Height Mod pixelSize)) As String
        Dim TextTable() As String = {"  ", "..", ",,", "::", ";;", "~~", "--", "++", "ii", "!!", "ll", "II", "??", "rr", "cc", "vv", "uu", "LL", "CC", "JJ", "UU", "YY", "XX", "ZZ", "00", "QQ", "WW", "MM", "BB", "88", "&&", "%%", "$$", "##", "@@"}

        For y As Integer = 0 To bm.Height - 1 Step pixelSize

            For x As Integer = 0 To bm.Width - 1 Step pixelSize

                brightness = bm.GetPixel(x, y).GetBrightness()
                position = (brightness / Convert.ToDouble(1 / TextTable.Length))
                asciiArt((x / pixelSize), (y / pixelSize)) = TextTable(position - 1)
            Next
        Next

        Return asciiArt
    End Function

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        pixelSize = Int(TrackBar1.Value)
    End Sub
End Class


Public Class Counter
    Private formOne As Form1
    Private threadnumber As Integer

    Public Sub New(ByVal my_form As Form1, ByVal my_number _
        As Integer)
        formOne = my_form
        threadNumber = my_number
    End Sub

    'http://www.vb-helper.com/howto_net_run_threads.html

    Public Sub Run()
        Try
            Do
                MsgBox("thread")
                Thread.Sleep(1000)

                SyncLock formOne


                End SyncLock
            Loop
        Catch ex As Exception
            Debug.WriteLine("Unexpected error in thread ")
        End Try
    End Sub


End Class
