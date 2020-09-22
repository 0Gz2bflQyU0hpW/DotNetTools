Option Infer Off
Option Explicit On
Option Strict On
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Collections
Public Class TiffManager
    Private _ImageFileName As String
    Private _PageNumber As Integer
    Private image As Image
    Private _TempWorkingDir As String
    Public Sub New(ByVal imageFileName As String)
        Me._ImageFileName = imageFileName
        image = image.FromFile(_ImageFileName)
        GetPageNumber()
    End Sub

    Public Sub New()
    End Sub

    ''' <summary>
    ''' Read the image file for the page number.
    ''' </summary>
    Private Sub GetPageNumber()
        Dim objGuid As Guid = image.FrameDimensionsList(0)
        Dim objDimension As New FrameDimension(objGuid)

        'Gets the total number of frames in the .tiff file
        _PageNumber = image.GetFrameCount(objDimension)

        Exit Sub
    End Sub

    ''' <summary>
    ''' Read the image base string,
    ''' Assert(GetFileNameStartString(@"c:\test\abc.tif"),"abc")
    ''' </summary>
    ''' <param name="strFullName"></param>
    ''' <returns></returns>
    Private Function GetFileNameStartString(ByVal strFullName As String) As String
        Dim posDot As Integer = _ImageFileName.LastIndexOf(".")
        Dim posSlash As Integer = _ImageFileName.LastIndexOf("\")
        Return _ImageFileName.Substring(posSlash + 1, posDot - posSlash - 1)
    End Function

    ''' <summary>
    ''' This function will output the image to a TIFF file with specific compression format
    ''' </summary>
    ''' <param name="outPutDirectory">The splited images' directory</param>
    ''' <param name="format">The codec for compressing</param>
    ''' <returns>splited file name array list</returns>
    Public Function SplitTiffImage(ByVal outPutDirectory As String, ByVal format As EncoderValue) As ArrayList
        Dim fileStartString As String = (outPutDirectory & "\") + GetFileNameStartString(_ImageFileName)
        Dim splitedFileNames As New ArrayList()
        Try
            Dim objGuid As Guid = image.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(objGuid)

            'Saves every frame as a separate file.
            Dim enc As Encoder = Encoder.Compression
            Dim curFrame As Integer = 0
            For i As Integer = 0 To _PageNumber - 1
                image.SelectActiveFrame(objDimension, curFrame)
                Dim ep As New EncoderParameters(1)
                ep.Param(0) = New EncoderParameter(enc, CLng(format))
                Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

                'Save the master bitmap
                Dim fileName As String = String.Format("{0}{1}.TIF", fileStartString, i.ToString())
                image.Save(fileName, info, ep)
                splitedFileNames.Add(fileName)

                curFrame += 1
            Next
        Catch generatedExceptionName As Exception
            Throw
        End Try

        Return splitedFileNames
    End Function

    ''' <summary>
    ''' This function will join the TIFF file with a specific compression format
    ''' </summary>
    ''' <param name="imageFiles">string array with source image files</param>
    ''' <param name="outFile">target TIFF file to be produced</param>
    ''' <param name="compressEncoder">compression codec enum</param>
    Public Sub JoinTiffImages(ByVal imageFiles As String(), ByVal outFile As String, ByVal compressEncoder As EncoderValue)
        Try
            'If only one page in the collection, copy it directly to the target file.
            If imageFiles.Length = 1 Then
                File.Copy(imageFiles(0), outFile, True)
                Exit Sub
            End If

            'use the save encoder
            Dim enc As Encoder = Encoder.SaveFlag

            Dim ep As New EncoderParameters(2)
            ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
            ep.Param(1) = New EncoderParameter(Encoder.Compression, CLng(compressEncoder))

            Dim pages As Bitmap = Nothing
            Dim frame As Integer = 0
            Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")


            For Each strImageFile As String In imageFiles
                If frame = 0 Then
                    pages = DirectCast(image.FromFile(strImageFile), Bitmap)

                    'save the first frame
                    pages.Save(outFile, info, ep)
                Else
                    'save the intermediate frames
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

                    Dim bm As Bitmap = DirectCast(image.FromFile(strImageFile), Bitmap)
                    pages.SaveAdd(bm, ep)
                End If

                If frame = imageFiles.Length - 1 Then
                    'flush and close.
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
                    pages.SaveAdd(ep)
                End If

                frame += 1
            Next
        Catch generatedExceptionName As Exception
            Throw
        End Try

        Exit Sub
    End Sub

    ''' <summary>
    ''' This function will join the TIFF file with a specific compression format
    ''' </summary>
    ''' <param name="imageFiles">array list with source image files</param>
    ''' <param name="outFile">target TIFF file to be produced</param>
    ''' <param name="compressEncoder">compression codec enum</param>
    Public Sub JoinTiffImages(ByVal imageFiles As ArrayList, ByVal outFile As String, ByVal compressEncoder As EncoderValue)
        Try
            'If only one page in the collection, copy it directly to the target file.
            If imageFiles.Count = 1 Then
                File.Copy(DirectCast(imageFiles(0), String), outFile, True)
                Exit Sub
            End If

            'use the save encoder
            Dim enc As Encoder = Encoder.SaveFlag

            Dim ep As New EncoderParameters(2)
            ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
            ep.Param(1) = New EncoderParameter(Encoder.Compression, CLng(compressEncoder))

            Dim pages As Bitmap = Nothing
            Dim frame As Integer = 0
            Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")


            For Each strImageFile As String In imageFiles
                If frame = 0 Then
                    pages = DirectCast(image.FromFile(strImageFile), Bitmap)

                    'save the first frame
                    pages.Save(outFile, info, ep)
                Else
                    'save the intermediate frames
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

                    Dim bm As Bitmap = DirectCast(image.FromFile(strImageFile), Bitmap)
                    pages.SaveAdd(bm, ep)
                    bm.Dispose()
                End If

                If frame = imageFiles.Count - 1 Then
                    'flush and close.
                    ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
                    pages.SaveAdd(ep)
                End If

                frame += 1
            Next
        Catch ex As Exception
#If DEBUG Then
            Console.WriteLine(ex.Message)
#End If
            Throw
        End Try

        Exit Sub
    End Sub

    ''' <summary>
    ''' Remove a specific page within the image object and save the result to an output image file.
    ''' </summary>
    ''' <param name="pageNumber">page number to be removed</param>
    ''' <param name="compressEncoder">compress encoder after operation</param>
    ''' <param name="strFileName">filename to be outputed</param>
    ''' <returns></</returns>
    Public Sub RemoveAPage(ByVal pageNumber As Integer, ByVal compressEncoder As EncoderValue, ByVal strFileName As String)
        Try
            'Split the image files to single pages.
            Dim arrSplited As ArrayList = SplitTiffImage(Me._TempWorkingDir, compressEncoder)

            'Remove the specific page from the collection
            Dim strPageRemove As String = String.Format("{0}\{1}{2}.TIF", _TempWorkingDir, GetFileNameStartString(Me._ImageFileName), pageNumber)
            arrSplited.Remove(strPageRemove)

            JoinTiffImages(arrSplited, strFileName, compressEncoder)
        Catch generatedExceptionName As Exception
            Throw
        End Try

        Exit Sub
    End Sub

    ''' <summary>
    ''' Getting the supported codec info.
    ''' </summary>
    ''' <param name="mimeType">description of mime type</param>
    ''' <returns>image codec info</returns>
    Private Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim encoders As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
        For j As Integer = 0 To encoders.Length - 1
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
        Next

        Throw New Exception(mimeType & " mime type not found in ImageCodecInfo")
    End Function

    ''' <summary>
    ''' Return the memory steam of a specific page
    ''' </summary>
    ''' <param name="pageNumber">page number to be extracted</param>
    ''' <returns>image object</returns>
    Public Function GetSpecificPage(ByVal pageNumber As Integer) As Image
        Dim ms As MemoryStream = Nothing
        Dim retImage As Image = Nothing
        Try
            ms = New MemoryStream()
            Dim objGuid As Guid = image.FrameDimensionsList(0)
            Dim objDimension As New FrameDimension(objGuid)

            image.SelectActiveFrame(objDimension, pageNumber)
            image.Save(ms, ImageFormat.Bmp)

            retImage = image.FromStream(ms)

            Return retImage
        Catch generatedExceptionName As Exception
            ms.Close()
            retImage.Dispose()
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Convert the existing TIFF to a different codec format
    ''' </summary>
    ''' <param name="compressEncoder"></param>
    ''' <returns></returns>
    Public Sub ConvertTiffFormat(ByVal strNewImageFileName As String, ByVal compressEncoder As EncoderValue)
        'Split the image files to single pages.
        Dim arrSplited As ArrayList = SplitTiffImage(Me._TempWorkingDir, compressEncoder)
        JoinTiffImages(arrSplited, strNewImageFileName, compressEncoder)

        Exit Sub
    End Sub

    ''' <summary>
    ''' Image file to operate
    ''' </summary>
    Public Property ImageFileName() As String
        Get
            Return _ImageFileName
        End Get
        Set(ByVal value As String)
            _ImageFileName = value
        End Set
    End Property

    ''' <summary>
    ''' Buffering directory
    ''' </summary>
    Public Property TempWorkingDir() As String
        Get
            Return _TempWorkingDir
        End Get
        Set(ByVal value As String)
            _TempWorkingDir = value
        End Set
    End Property

    ''' <summary>
    ''' Image page number
    ''' </summary>
    Public ReadOnly Property PageNumber() As Integer
        Get
            Return _PageNumber
        End Get
    End Property


#Region "IDisposable Members"

    Public Sub Dispose()
        image.Dispose()
        System.GC.SuppressFinalize(Me)
    End Sub

#End Region
End Class
