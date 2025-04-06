Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports WicNet

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim filename = "240509_EXPO20th_HPbnr_1000-480.jpg"

        Me.Text = filename
        WicLoadFromFile(filename)
    End Sub

    Private Sub WicLoadFromFile(ByVal value As String)
        Dim src = WicBitmapSource.Load(value)
        Dim bmp = ConvertWicBitmapSourceToBitmap(src)

        PictureBox1.Image = bmp
        PictureBox1.Size = PictureBox1.Image.Size
    End Sub

    Private Function ConvertWicBitmapSourceToBitmap(wicSource As WicBitmapSource) As Bitmap
        If wicSource Is Nothing Then
            Return Nothing
        End If

        Dim width As Integer = wicSource.Size.Width
        Dim height As Integer = wicSource.Size.Height
        Dim pixelFormat As PixelFormat = GetPixelFormat(wicSource.PixelFormat)

        Dim bitmap As New Bitmap(width, height, pixelFormat)

        Dim bitmapData As BitmapData = bitmap.LockBits(
            New Rectangle(0, 0, width, height),
            ImageLockMode.WriteOnly, pixelFormat)

        Dim stride As Integer = bitmapData.Stride
        Dim bufferSize As UInteger = CUInt(stride * height)

        Try
            wicSource.CopyPixels(bufferSize, bitmapData.Scan0, CUInt(stride))
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            bitmap.UnlockBits(bitmapData)
        End Try

        Return bitmap
    End Function

    Private Function GetPixelFormat(ByVal wicFormat As Guid) As PixelFormat
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormatDontCare) Then
            Debug.WriteLine("!")
            Return PixelFormat.DontCare
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat32bppPRGBA) Then
            Debug.WriteLine("!")
            Return PixelFormat.Format16bppArgb1555
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat32bppBGR) Then
            Return PixelFormat.Format32bppRgb
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat32bppBGRA) Then
            Return PixelFormat.Format32bppArgb
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat24bppBGR) Then
            Return PixelFormat.Format24bppRgb
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat24bppRGB) Then
            Return PixelFormat.Format24bppRgb
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat8bppIndexed) Then
            Return PixelFormat.Format8bppIndexed
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat4bppIndexed) Then
            Return PixelFormat.Format4bppIndexed
        End If
        If wicFormat.Equals(WicPixelFormat.GUID_WICPixelFormat1bppIndexed) Then
            Return PixelFormat.Format1bppIndexed
        End If

        Return PixelFormat.Format32bppArgb
    End Function

End Class
