Imports System.Drawing
Public Class SpinningProgress
    Private m_InactiveColour As Color = Color.FromArgb(218, 218, 218)
    Private m_ActiveColour As Color = Color.FromArgb(35, 146, 33)
    Private m_TransistionColour As Color = Color.FromArgb(129, 242, 121)

    Private innerBackgroundRegion As Region
    Private segmentPaths(11) As Drawing2D.GraphicsPath

    Private m_AutoIncrement As Boolean = True
    Private m_IncrementFrequency As Double = 100
    Private m_BehindIsActive As Boolean = True
    Private m_TransitionSegment As Integer = 0

    Private m_AutoRotateTimer As System.Timers.Timer

    <System.ComponentModel.DefaultValue(GetType(Color), "218, 218, 218")> _
    Property InactiveSegmentColour() As Color
        Get
            Return m_InactiveColour
        End Get
        Set(ByVal value As Color)
            m_InactiveColour = value
            Invalidate()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Color), "35, 146, 33")> _
    Property ActiveSegmentColour() As Color
        Get
            Return m_ActiveColour
        End Get
        Set(ByVal value As Color)
            m_ActiveColour = value
            Invalidate()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(GetType(Color), "129, 242, 121")> _
    Property TransistionSegmentColour() As Color
        Get
            Return m_TransistionColour
        End Get
        Set(ByVal value As Color)
            m_TransistionColour = value
            Invalidate()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(True)> _
    Property BehindTransistionSegmentIsActive() As Boolean
        Get
            Return m_BehindIsActive
        End Get
        Set(ByVal value As Boolean)
            m_BehindIsActive = value
            Invalidate()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(-1)> _
    Property TransistionSegment() As Integer
        Get
            Return m_TransitionSegment
        End Get
        Set(ByVal value As Integer)
            If value > 11 Or value < -1 Then
                Throw New ArgumentException("TransistionSegment must be between -1 and 11")
            End If
            m_TransitionSegment = value
            Invalidate()
        End Set
    End Property

    <System.ComponentModel.DefaultValue(True)> _
    Property AutoIncrement() As Boolean
        Get
            Return m_AutoIncrement
        End Get
        Set(ByVal value As Boolean)
            m_AutoIncrement = value

            If value = False AndAlso m_AutoRotateTimer IsNot Nothing Then
                m_AutoRotateTimer.Dispose()
                m_AutoRotateTimer = Nothing
            End If

            If value = True AndAlso m_AutoRotateTimer Is Nothing Then
                m_AutoRotateTimer = New System.Timers.Timer(m_IncrementFrequency)
                AddHandler m_AutoRotateTimer.Elapsed, AddressOf IncrementTransisionSegment
                m_AutoRotateTimer.Start()
            End If
        End Set
    End Property

    <System.ComponentModel.DefaultValue(CDbl(100))> _
    Property AutoIncrementFrequency() As Double
        Get
            Return m_IncrementFrequency
        End Get
        Set(ByVal value As Double)
            m_IncrementFrequency = value

            If m_AutoRotateTimer IsNot Nothing Then
                AutoIncrement = False
                AutoIncrement = True
            End If
        End Set
    End Property

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CalculateSegments()

        m_AutoRotateTimer = New System.Timers.Timer(m_IncrementFrequency)
        AddHandler m_AutoRotateTimer.Elapsed, AddressOf IncrementTransisionSegment
        m_AutoRotateTimer.Start()
    End Sub

    Private Sub IncrementTransisionSegment(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        If m_TransitionSegment = 11 Then
            m_TransitionSegment = 0
            m_BehindIsActive = Not m_BehindIsActive
        ElseIf m_TransitionSegment = -1 Then
            m_TransitionSegment = 0
        Else
            m_TransitionSegment += 1
        End If

        Invalidate()
    End Sub

    Private Sub CalculateSegments()
        Dim rctFull As New Rectangle(0, _
                                     0, _
                                     Me.Width, _
                                     Me.Height)
        Dim rctInner As New Rectangle(CInt(Me.Width * (7 / 30)), _
                                      CInt(Me.Height * (7 / 30)), _
                                      CInt(Me.Width - (Me.Width * (7 / 30) * 2)), _
                                      CInt(Me.Height - (Me.Height * (7 / 30) * 2)))
        Dim pthInnerBackground As Drawing2D.GraphicsPath

        'Create 12 segment pieces
        For intCount As Integer = 0 To 11
            segmentPaths(intCount) = New Drawing2D.GraphicsPath

            'We subtract 90 so that the starting segment is at 12 o'clock
            segmentPaths(intCount).AddPie(rctFull, (intCount * 30) - 90, 25)
        Next

        'Create the center circle cut-out
        pthInnerBackground = New Drawing2D.GraphicsPath
        pthInnerBackground.AddPie(rctInner, 0, 360)
        innerBackgroundRegion = New Region(pthInnerBackground)
    End Sub

    Private Sub SpinningProgress_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.EnabledChanged
        If Enabled = True Then
            If m_AutoRotateTimer IsNot Nothing Then
                m_AutoRotateTimer.Start()
            End If
        Else
            If m_AutoRotateTimer IsNot Nothing Then
                m_AutoRotateTimer.Stop()
            End If
        End If
    End Sub

    Private Sub ProgressDisk_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        e.Graphics.ExcludeClip(innerBackgroundRegion)

        For intCount As Integer = 0 To 11
            If Me.Enabled Then
                If intCount = m_TransitionSegment Then
                    'If this segment is the transistion segment, colour it differently
                    e.Graphics.FillPath(New SolidBrush(m_TransistionColour), segmentPaths(intCount))
                ElseIf intCount < m_TransitionSegment Then
                    'This segment is behind the transistion segment
                    If m_BehindIsActive Then
                        'If behind the transistion should be active, 
                        'colour it with the active colour
                        e.Graphics.FillPath(New SolidBrush(m_ActiveColour), segmentPaths(intCount))
                    Else
                        'If behind the transistion should be in-active, 
                        'colour it with the in-active colour
                        e.Graphics.FillPath(New SolidBrush(m_InactiveColour), segmentPaths(intCount))
                    End If
                Else
                    'This segment is ahead of the transistion segment
                    If m_BehindIsActive Then
                        'If behind the the transistion should be active, 
                        'colour it with the in-active colour
                        e.Graphics.FillPath(New SolidBrush(m_InactiveColour), segmentPaths(intCount))
                    Else
                        'If behind the the transistion should be in-active, 
                        'colour it with the active colour
                        e.Graphics.FillPath(New SolidBrush(m_ActiveColour), segmentPaths(intCount))
                    End If
                End If
            Else
                'Draw all segments in in-active colour if not enabled
                e.Graphics.FillPath(New SolidBrush(m_InactiveColour), segmentPaths(intCount))
            End If
        Next
    End Sub

    Private Sub ProgressDisk_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        CalculateSegments()
    End Sub

    Private Sub ProgressDisk_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        CalculateSegments()
    End Sub
End Class
