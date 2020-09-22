Option Infer Off
Option Explicit On
Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Collections
Imports System.Threading
Imports System.Drawing.Imaging
Imports System.IO
Public Class TiffUtil
#Region "Variable & Class Definitions"

    Private Shared tifImageCodecInfo As System.Drawing.Imaging.ImageCodecInfo

    Private Shared tifEncoderParameterMultiFrame As EncoderParameter
    Private Shared tifEncoderParameterFrameDimensionPage As EncoderParameter
    Private Shared tifEncoderParameterFlush As EncoderParameter
    Private Shared tifEncoderParameterCompression As EncoderParameter
    Private Shared tifEncoderParameterLastFrame As EncoderParameter
    Private Shared tifEncoderParameter24BPP As EncoderParameter
    Private Shared tifEncoderParameter1BPP As EncoderParameter

    Private Shared tifEncoderParametersPage1 As EncoderParameters
    Private Shared tifEncoderParametersPageX As EncoderParameters
    Private Shared tifEncoderParametersPageLast As EncoderParameters

    Private Shared tifEncoderSaveFlag As System.Drawing.Imaging.Encoder
    Private Shared tifEncoderCompression As System.Drawing.Imaging.Encoder
    Private Shared tifEncoderColorDepth As System.Drawing.Imaging.Encoder

    Private Shared encoderAssigned As Boolean

    Public Shared tempDir As String
    Public Shared initComplete As Boolean

    Public Sub New(ByVal tempPath As String)
        Try
            If Not initComplete Then
                If Not tempPath.EndsWith("\") Then
                    tempDir = tempPath & "\"
                Else
                    tempDir = tempPath
                End If

                Directory.CreateDirectory(tempDir)
                initComplete = True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Throw ex
        End Try
    End Sub

#End Region

#Region "Retrieve Page Count of a multi-page TIFF file"

    Public Function getPageCount(ByVal fileName As String) As Integer
        Dim pageCount As Integer = -1

        Try
            Dim img As Image = Bitmap.FromFile(fileName)
            pageCount = img.GetFrameCount(FrameDimension.Page)
            img.Dispose()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return pageCount
    End Function

    Public Function getPageCount(ByVal img As Image) As Integer
        Dim pageCount As Integer = -1
        Try
            pageCount = img.GetFrameCount(FrameDimension.Page)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Return pageCount
    End Function

#End Region

#Region "Retrieve multiple single page images from a single multi-page TIFF file"

    Public Function getTiffImages(ByVal sourceImage As Image, ByVal pageNumbers As Integer()) As Image()
        Dim ms As MemoryStream = Nothing
        Dim returnImage As Image() = New Image(pageNumbers.Length - 1) {}

        Try
            Dim objGuid As Guid = sourceImage.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(objGuid)

            For i As Integer = 0 To pageNumbers.Length - 1
                ms = New MemoryStream()
                sourceImage.SelectActiveFrame(objDimension, pageNumbers(i))
                sourceImage.Save(ms, ImageFormat.Tiff)
                returnImage(i) = Image.FromStream(ms)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ms.Close()
        End Try

        Return returnImage
    End Function

    Public Function getTiffImages(ByVal sourceImage As Image) As Image()
        Dim ms As MemoryStream = Nothing
        Dim pageCount As Integer = getPageCount(sourceImage)

        Dim returnImage As Image() = New Image(pageCount - 1) {}

        Try
            Dim objGuid As Guid = sourceImage.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(objGuid)

            For i As Integer = 0 To pageCount - 1
                ms = New MemoryStream()
                sourceImage.SelectActiveFrame(objDimension, i)
                sourceImage.Save(ms, ImageFormat.Tiff)
                returnImage(i) = Image.FromStream(ms)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ms.Close()
        End Try

        Return returnImage
    End Function

    Public Function getTiffImages(ByVal sourceFile As String, ByVal pageNumbers As Integer()) As Image()
        Dim returnImage As Image() = New Image(pageNumbers.Length - 1) {}

        Try
            Dim sourceImage As Image = Bitmap.FromFile(sourceFile)
            returnImage = getTiffImages(sourceImage, pageNumbers)
            sourceImage.Dispose()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            returnImage = Nothing
        End Try

        Return returnImage
    End Function

#End Region

#Region "Retrieve a specific page from a multi-page TIFF image"

    Public Function getTiffImage(ByVal sourceFile As String, ByVal pageNumber As Integer) As Image
        Dim returnImage As Image = Nothing

        Try
            Dim sourceImage As Image = Image.FromFile(sourceFile)
            returnImage = getTiffImage(sourceImage, pageNumber)
            sourceImage.Dispose()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            returnImage = Nothing
        End Try

        Return returnImage
    End Function

    Public Function getTiffImage(ByVal sourceImage As Image, ByVal pageNumber As Integer) As Image
        Dim ms As MemoryStream = Nothing
        Dim returnImage As Image = Nothing

        Try
            ms = New MemoryStream()
            Dim objGuid As Guid = sourceImage.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(objGuid)
            sourceImage.SelectActiveFrame(objDimension, pageNumber)
            sourceImage.Save(ms, ImageFormat.Tiff)
            returnImage = Image.FromStream(ms)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            ms.Close()
        End Try

        Return returnImage
    End Function

    Public Function getTiffImage(ByVal sourceFile As String, ByVal targetFile As String, ByVal pageNumber As Integer) As Boolean
        Dim response As Boolean = False

        Try
            Dim returnImage As Image = getTiffImage(sourceFile, pageNumber)
            returnImage.Save(targetFile)
            returnImage.Dispose()
            response = True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return response
    End Function

#End Region

#Region "Split a multi-page TIFF file into multiple single page TIFF files"

    Public Function SplitTiffPages(ByVal SourceFile As String, ByVal TargetDirectory As String) As String()
        Dim returnImages As String()

        Try
            Dim sourceImage As Image = Bitmap.FromFile(SourceFile)
            Dim sourceImages As Image() = splitTiffPages(sourceImage)

            Dim pageCount As Integer = sourceImages.Length

            returnImages = New String(pageCount - 1) {}
            For i As Integer = 0 To pageCount - 1
                Dim fi As New FileInfo(SourceFile)
                Dim babyImg As String = ((TargetDirectory & "\") + fi.Name.Substring(0, (fi.Name.Length - fi.Extension.Length)) & "_PAGE") + (i + 1).ToString().PadLeft(3, "0"c) + fi.Extension
                sourceImages(i).Save(babyImg)
                returnImages(i) = babyImg
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            returnImages = Nothing
        End Try
        Return returnImages
    End Function

    Public Function SplitTiffPages(ByVal sourceImage As Image) As Image()
        Dim returnImages As Image()

        Try
            Dim pageCount As Integer = getPageCount(sourceImage)
            returnImages = New Image(pageCount - 1) {}

            For i As Integer = 0 To pageCount - 1
                Dim img As Image = getTiffImage(sourceImage, i)
                returnImages(i) = DirectCast(img.Clone(), Image)
                img.Dispose()

            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            returnImages = Nothing
        End Try

        Return returnImages
    End Function

#End Region

#Region "Merge multiple single page TIFF to a single multi page TIFF"

    Public Function mergeTiffPages(ByVal sourceFiles As String(), ByVal targetFile As String) As Boolean
        Dim response As Boolean = False
        Try
            assignEncoder()
            ' If only 1 page was passed, copy directly to output
            If sourceFiles.Length = 1 Then
                File.Copy(sourceFiles(0), targetFile, True)
                Return True
            End If
            Dim pageCount As Integer = sourceFiles.Length
            ' First page
            Dim finalImage As Image = Image.FromFile(sourceFiles(0))
            finalImage.Save(targetFile, tifImageCodecInfo, tifEncoderParametersPage1)
            ' All other pages
            For i As Integer = 1 To pageCount - 1
                Dim img As Image = Image.FromFile(sourceFiles(i))
                finalImage.SaveAdd(img, tifEncoderParametersPageX)
                img.Dispose()
            Next
            ' Last page
            finalImage.SaveAdd(tifEncoderParametersPageLast)
            finalImage.Dispose()
            response = True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            response = False
        End Try

        Return response
    End Function

    Public Function mergeTiffPages(ByVal sourceFile As String, ByVal targetFile As String, ByVal pageNumbers As Integer()) As Boolean
        Dim response As Boolean = False

        Try
            assignEncoder()

            ' Get individual Images from the original image 
            Dim sourceImage As Image = Bitmap.FromFile(sourceFile)
            Dim ms As New MemoryStream()
            Dim sourceImages As Image() = New Image(pageNumbers.Length - 1) {}
            Dim guid As Guid = sourceImage.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(guid)
            For i As Integer = 0 To pageNumbers.Length - 1
                sourceImage.SelectActiveFrame(objDimension, pageNumbers(i))
                sourceImage.Save(ms, ImageFormat.Tiff)
                sourceImages(i) = Image.FromStream(ms)
            Next

            ' Merge individual Images into one Image 
            ' First page
            Dim finalImage As Image = sourceImages(0)
            finalImage.Save(targetFile, tifImageCodecInfo, tifEncoderParametersPage1)
            ' All other pages
            For i As Integer = 1 To pageNumbers.Length - 1
                finalImage.SaveAdd(sourceImages(i), tifEncoderParametersPageX)
            Next
            ' Last page
            finalImage.SaveAdd(tifEncoderParametersPageLast)
            finalImage.Dispose()

            response = True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return response
    End Function

    Public Function mergeTiffPagesAlternate(ByVal sourceFile As String, ByVal targetFile As String, ByVal pageNumbers As Integer()) As Boolean
        Dim response As Boolean = False

        Try
            ' Initialize the encoders, occurs once for the lifetime of the class
            assignEncoder()

            ' Get individual Images from the original image 
            Dim sourceImage As Image = Bitmap.FromFile(sourceFile)
            Dim msArray As MemoryStream() = New MemoryStream(pageNumbers.Length - 1) {}
            Dim guid As Guid = sourceImage.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(guid)
            For i As Integer = 0 To pageNumbers.Length - 1
                msArray(i) = New MemoryStream()
                sourceImage.SelectActiveFrame(objDimension, pageNumbers(i))
                sourceImage.Save(msArray(i), ImageFormat.Tiff)
            Next

            ' Merge individual page streams into single stream
            Dim ms As MemoryStream = mergeTiffStreams(msArray)
            Dim targetImage As Image = Bitmap.FromStream(ms)
            targetImage.Save(targetFile)

            response = True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return response
    End Function

    Public Function mergeTiffStreams(ByVal tifsStream As System.IO.MemoryStream()) As System.IO.MemoryStream
        Dim ep As EncoderParameters = Nothing
        Dim singleStream As New System.IO.MemoryStream()

        Try
            assignEncoder()

            Dim imgTif As Image = Image.FromStream(tifsStream(0))

            If tifsStream.Length > 1 Then
                ' Multi-Frame
                ep = New EncoderParameters(2)
                ep.Param(0) = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.MultiFrame))
                ep.Param(1) = New EncoderParameter(tifEncoderCompression, CLng(EncoderValue.CompressionRle))
            Else
                ' Single-Frame
                ep = New EncoderParameters(1)
                ep.Param(0) = New EncoderParameter(tifEncoderCompression, CLng(EncoderValue.CompressionRle))
            End If

            'Save the first page
            imgTif.Save(singleStream, tifImageCodecInfo, ep)

            If tifsStream.Length > 1 Then
                ep = New EncoderParameters(2)
                ep.Param(0) = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.FrameDimensionPage))

                'Add the rest of pages
                For i As Integer = 1 To tifsStream.Length - 1
                    Dim pgTif As Image = Image.FromStream(tifsStream(i))

                    ep.Param(1) = New EncoderParameter(tifEncoderCompression, CLng(EncoderValue.CompressionRle))

                    imgTif.SaveAdd(pgTif, ep)
                Next

                ep = New EncoderParameters(1)
                ep.Param(0) = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.Flush))
                imgTif.SaveAdd(ep)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If ep IsNot Nothing Then
                ep.Dispose()
            End If
        End Try

        Return singleStream
    End Function

#End Region

#Region "Internal support functions"

    Private Sub assignEncoder()
        Try
            If encoderAssigned = True Then
                Exit Sub
            End If

            For Each ici As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
                If ici.MimeType = "image/tiff" Then
                    tifImageCodecInfo = ici
                End If
            Next

            tifEncoderSaveFlag = System.Drawing.Imaging.Encoder.SaveFlag
            tifEncoderCompression = System.Drawing.Imaging.Encoder.Compression
            tifEncoderColorDepth = System.Drawing.Imaging.Encoder.ColorDepth

            tifEncoderParameterMultiFrame = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.MultiFrame))
            tifEncoderParameterFrameDimensionPage = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.FrameDimensionPage))
            tifEncoderParameterFlush = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.Flush))
            tifEncoderParameterCompression = New EncoderParameter(tifEncoderCompression, CLng(EncoderValue.CompressionRle))
            tifEncoderParameterLastFrame = New EncoderParameter(tifEncoderSaveFlag, CLng(EncoderValue.LastFrame))
            tifEncoderParameter24BPP = New EncoderParameter(tifEncoderColorDepth, CLng(24))
            tifEncoderParameter1BPP = New EncoderParameter(tifEncoderColorDepth, CLng(8))

            ' ******************************************************************* //
            ' *** Have only 1 of the following 3 groups assigned for encoders *** //
            ' ******************************************************************* //

            ' Regular 
            tifEncoderParametersPage1 = New EncoderParameters(1)
            tifEncoderParametersPage1.Param(0) = tifEncoderParameterMultiFrame
            tifEncoderParametersPageX = New EncoderParameters(1)
            tifEncoderParametersPageX.Param(0) = tifEncoderParameterFrameDimensionPage
            tifEncoderParametersPageLast = New EncoderParameters(1)
            tifEncoderParametersPageLast.Param(0) = tifEncoderParameterFlush

            '''/ Regular 
            'tifEncoderParametersPage1 = new EncoderParameters(2); 
            'tifEncoderParametersPage1.Param[0] = tifEncoderParameterMultiFrame;
            'tifEncoderParametersPage1.Param[1] = tifEncoderParameterCompression;
            'tifEncoderParametersPageX = new EncoderParameters(2); 
            'tifEncoderParametersPageX.Param[0] = tifEncoderParameterFrameDimensionPage; 
            'tifEncoderParametersPageX.Param[1] = tifEncoderParameterCompression;
            'tifEncoderParametersPageLast = new EncoderParameters(2); 
            'tifEncoderParametersPageLast.Param[0] = tifEncoderParameterFlush;
            'tifEncoderParametersPageLast.Param[1] = tifEncoderParameterLastFrame;

            '''/ 24 BPP Color 
            'tifEncoderParametersPage1 = new EncoderParameters(2); 
            'tifEncoderParametersPage1.Param[0] = tifEncoderParameterMultiFrame;
            'tifEncoderParametersPage1.Param[1] = tifEncoderParameter24BPP;
            'tifEncoderParametersPageX = new EncoderParameters(2); 
            'tifEncoderParametersPageX.Param[0] = tifEncoderParameterFrameDimensionPage;
            'tifEncoderParametersPageX.Param[1] = tifEncoderParameter24BPP;
            'tifEncoderParametersPageLast = new EncoderParameters(2); 
            'tifEncoderParametersPageLast.Param[0] = tifEncoderParameterFlush;
            'tifEncoderParametersPageLast.Param[1] = tifEncoderParameterLastFrame;

            '''/ 1 BPP BW 
            'tifEncoderParametersPage1 = new EncoderParameters(2); 
            'tifEncoderParametersPage1.Param[0] = tifEncoderParameterMultiFrame;
            'tifEncoderParametersPage1.Param[1] = tifEncoderParameterCompression;
            'tifEncoderParametersPageX = new EncoderParameters(2); 
            'tifEncoderParametersPageX.Param[0] = tifEncoderParameterFrameDimensionPage;
            'tifEncoderParametersPageX.Param[1] = tifEncoderParameterCompression;
            'tifEncoderParametersPageLast = new EncoderParameters(2); 
            'tifEncoderParametersPageLast.Param[0] = tifEncoderParameterFlush;
            'tifEncoderParametersPageLast.Param[1] = tifEncoderParameterLastFrame;

            encoderAssigned = True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Throw ex
        End Try
    End Sub

    Private Function ConvertToGrayscale(ByVal source As Bitmap) As Bitmap
        Try
            Dim bm As New Bitmap(source.Width, source.Height)
            Dim g As Graphics = Graphics.FromImage(bm)

            Dim cm As New ColorMatrix(New Single()() {New Single() {0.5F, 0.5F, 0.5F, 0, 0}, New Single() {0.5F, 0.5F, 0.5F, 0, 0}, New Single() {0.5F, 0.5F, 0.5F, 0, 0}, New Single() {0, 0, 0, 1, 0, 0}, New Single() {0, 0, 0, 0, 1, 0}, New Single() {0, 0, 0, 0, 0, 1}})
            Dim ia As New ImageAttributes()
            ia.SetColorMatrix(cm)
            g.DrawImage(source, New Rectangle(0, 0, source.Width, source.Height), 0, 0, source.Width, source.Height, _
            GraphicsUnit.Pixel, ia)
            g.Dispose()

            Return bm
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Throw ex
        End Try
    End Function

#End Region
#Region "ReducirTif"
    Public Sub ReducirTif(ByVal OriginalFile As String, ByVal NewFile As String, ByVal NewWidth As Integer, ByVal MaxHeight As Integer, ByVal OnlyResizeIfWider As Boolean)
        Dim FullsizeImage As System.Drawing.Image = System.Drawing.Image.FromFile(OriginalFile)
        ' Prevent using images internal thumbnail
        FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        If OnlyResizeIfWider Then
            If FullsizeImage.Width <= NewWidth Then
                NewWidth = FullsizeImage.Width
            End If
        End If
        Dim NewHeight As Integer = CInt(FullsizeImage.Height * NewWidth / FullsizeImage.Width)
        If NewHeight > MaxHeight Then
            ' Resize with height instead
            NewWidth = CInt(FullsizeImage.Width * MaxHeight / FullsizeImage.Height)
            NewHeight = MaxHeight
        End If
        Dim NewImage As System.Drawing.Image = FullsizeImage.GetThumbnailImage(NewWidth, NewHeight, Nothing, IntPtr.Zero)
        ' Clear handle to original file so that we can overwrite it if necessary
        FullsizeImage.Dispose()
        ' Save resized picture
        NewImage.Save(NewFile)
    End Sub
    Public Sub ReducirTif(ByVal OriginalFile As String, ByVal NewFile As String, ByVal PorcentajeFinal As Integer)
        Dim FullsizeImage As System.Drawing.Image = System.Drawing.Image.FromFile(OriginalFile)
        Dim NewWidth As Integer = CInt((FullsizeImage.Width * PorcentajeFinal) / 100)
        Dim MaxHeight As Integer = CInt((FullsizeImage.Height * PorcentajeFinal) / 100)
        ' Prevent using images internal thumbnail
        FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        If FullsizeImage.Width <= NewWidth Then
            NewWidth = FullsizeImage.Width
        End If
        Dim NewHeight As Integer = CInt(FullsizeImage.Height * NewWidth / FullsizeImage.Width)
        If NewHeight > MaxHeight Then
            ' Resize with height instead
            NewWidth = CInt(FullsizeImage.Width * MaxHeight / FullsizeImage.Height)
            NewHeight = MaxHeight
        End If
        Dim NewImage As System.Drawing.Image = FullsizeImage.GetThumbnailImage(NewWidth, NewHeight, Nothing, IntPtr.Zero)
        ' Clear handle to original file so that we can overwrite it if necessary
        FullsizeImage.Dispose()
        ' Save resized picture
        If NewFile.Trim.ToUpper.EndsWith(".PNG") Then
            NewImage.Save(NewFile, ImageFormat.Png)
        Else
            NewImage.Save(NewFile)
        End If
    End Sub
#End Region
End Class
