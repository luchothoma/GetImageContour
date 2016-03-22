Imports System.Math
Imports System.Drawing
Imports System.Drawing.Imaging

Module Module1

    Public Function ImgContorno(ByVal img As Bitmap, ByVal tolerancia As Integer) As Bitmap
        Dim img2 As New Bitmap(img.Width, img.Height)
        Dim lamatriz(img.Width, img.Height) As Integer
        Dim i As Integer, j As Integer
        Dim c As Color

        For j = 0 To img.Height - 1
            For i = 0 To img.Width - 1
                c = img.GetPixel(i, j)
                lamatriz(i, j) = (c.R * 70 + c.G * 150 + c.B * 29) / 256
            Next
        Next

        For j = 1 To img.Height - 1
            For i = 1 To img.Width - 1
                If Abs(lamatriz(i, j) - lamatriz(i, j - 1)) > tolerancia Or Abs(lamatriz(i, j) - lamatriz(i - 1, j)) > tolerancia Then
                    img2.SetPixel(i, j, Color.Black)
                Else
                    img2.SetPixel(i, j, Color.White)
                End If
            Next
        Next

        Return img2

    End Function

    Public Function MakeGrayscale(ByVal original As Bitmap) As Bitmap
        Dim i As Integer, j As Integer
        'make an empty bitmap the same size as original
        Dim newBitmap As Bitmap = New Bitmap(original.Width, original.Height)

        For i = 0 To original.Width - 1

            For j = 0 To original.Height - 1

                'get the pixel from the original image
                Dim originalColor As Color = original.GetPixel(i, j)

                'create the grayscale version of the pixel
                Dim grayScale As Integer = CInt((originalColor.R * 0.3) + (originalColor.G * 0.59) + (originalColor.B * 0.11))

                'create the color object
                Dim newColor As Color = Color.FromArgb(254, grayScale, grayScale, grayScale)
                'Dim newColor As Color = Color.FromArgb(254, CInt(originalColor.R * 0.3), CInt(originalColor.G * 0.4), CInt(originalColor.B * 0.2))

                'set the new image's pixel to the grayscale version
                newBitmap.SetPixel(i, j, newColor)
            Next
        Next

        Return newBitmap
    End Function

    Public Function MakeGrayscaleFast(ByVal img As Bitmap) As Bitmap
        Dim bm As Bitmap = New Bitmap(img.Width, img.Height)
        Dim g As Graphics = Graphics.FromImage(bm)
        Dim cm As ColorMatrix = New ColorMatrix(New Single()() _
             {New Single() {0.3, 0.3, 0.3, 0, 0}, _
            New Single() {0.59, 0.59, 0.59, 0, 0}, _
            New Single() {0.11, 0.11, 0.11, 0, 0}, _
            New Single() {0, 0, 0, 1, 0}, _
            New Single() {0, 0, 0, 0, 1}})

        Dim ia As ImageAttributes = New ImageAttributes()
        ia.SetColorMatrix(cm)
        g.DrawImage(img, New Rectangle(0, 0, img.Width, img.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia)
        g.Dispose()
        Return bm

    End Function

    Public Function Otsus(ByVal grey As Bitmap, ByVal level As Byte) As Bitmap
        Dim i As Integer, j As Integer
        Dim cantp As ULong = grey.Width * grey.Height
        Dim promp As ULong
        Dim lmed As Int16
        Dim newBitmap As Bitmap = New Bitmap(grey.Width, grey.Height)
        Dim col As Color
        Dim chek As Integer

        For i = 0 To grey.Width - 1

            For j = 0 To grey.Height - 1
                col = grey.GetPixel(i, j)
                promp = promp + CInt(col.R)
            Next
        Next

        lmed = CByte((promp / cantp) / 6 * (level))

        For i = 0 To grey.Width - 1

            For j = 0 To grey.Height - 1
                col = grey.GetPixel(i, j)
                chek = col.R
                If chek > lmed Then
                    newBitmap.SetPixel(i, j, Color.White)
                ElseIf chek < lmed Then
                    newBitmap.SetPixel(i, j, Color.Black)
                    'Else
                    ' newBitmap.SetPixel(i, j, Color.Red)
                End If
            Next
        Next

        Return newBitmap
    End Function

    Public Function ContornOtsus(ByVal bin1 As Bitmap, ByVal bin2 As Bitmap) As Bitmap
        Dim i As Integer, j As Integer
        Dim newBitmap As Bitmap = New Bitmap(bin1.Width, bin1.Height)
        Dim col1, col2 As Color


        For i = 0 To bin1.Width - 1

            For j = 0 To bin1.Height - 1
                col1 = bin1.GetPixel(i, j)
                col2 = bin2.GetPixel(i, j)
                ' If col1 = col2 And col1 = Color.Red Then
                'newBitmap.SetPixel(i, j, Color.Orange)
                'Else
                If col1 = col2 Then
                    If col1 = Color.FromArgb(255, 0, 0, 0) Then
                        newBitmap.SetPixel(i, j, Color.Black)
                    Else
                        newBitmap.SetPixel(i, j, Color.White)
                    End If
                Else
                    newBitmap.SetPixel(i, j, Color.Red)
                End If
            Next
        Next

        Return newBitmap
    End Function

    Public Function ContornOtsusBordes(ByVal img As Bitmap) As Bitmap
        Dim i As Integer, j As Integer
        Dim newBitmap As Bitmap = New Bitmap(img.Width, img.Height)
        Dim col, col1, col2, col3, col4 As Color


        For i = 1 To img.Width - 2

            For j = 1 To img.Height - 2
                col = img.GetPixel(i, j)
                col1 = img.GetPixel(i - 1, j)
                col2 = img.GetPixel(i + 1, j)
                col3 = img.GetPixel(i, j + 1)
                col4 = img.GetPixel(i, j - 1)
                If col <> col1 Or col <> col2 Or col <> col3 Or col <> col4 Then
                    newBitmap.SetPixel(i, j, Color.Black)
                End If
            Next
        Next

        Return newBitmap
    End Function

    Public Function ContornO(img As Bitmap) As Bitmap

        Dim grey As Bitmap = MakeGrayscaleFast(img)
        Dim i As Integer, j As Integer
        Dim cantp As ULong = grey.Width * grey.Height
        Dim promp As ULong = 0
        Dim lmed1, lmed2 As Int16
        Dim col, col1, col2, col3, col4 As Color
        Dim chek As Integer
        Dim final As Bitmap = New Bitmap(img.Width, img.Height)

        For i = 0 To grey.Width - 1

            For j = 0 To grey.Height - 1
                col = grey.GetPixel(i, j)
                promp = promp + CInt(col.R)
            Next
        Next

        lmed1 = CByte((promp / cantp) / 6 * (7))
        lmed2 = CByte((promp / cantp) / 6 * (6))

        For i = 0 To grey.Width - 1
            For j = 0 To grey.Height - 1
                col = grey.GetPixel(i, j)
                chek = col.R
                If chek >= lmed1 And chek >= lmed2 Then
                    grey.SetPixel(i, j, Color.White)
                ElseIf chek < lmed1 And chek < lmed2 Then
                    grey.SetPixel(i, j, Color.Black)
                Else
                    grey.SetPixel(i, j, Color.Red)
                End If
            Next
        Next

        For i = 1 To grey.Width - 2
            For j = 1 To grey.Height - 2
                col = grey.GetPixel(i, j)
                col1 = grey.GetPixel(i - 1, j)
                col2 = grey.GetPixel(i + 1, j)
                col3 = grey.GetPixel(i, j + 1)
                col4 = grey.GetPixel(i, j - 1)
                If col <> col1 Or col <> col2 Or col <> col3 Or col <> col4 Then
                    final.SetPixel(i, j, Color.Black)
                End If

            Next
        Next


        Return final

    End Function

End Module
