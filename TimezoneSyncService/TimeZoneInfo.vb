'Author: Arman Ghazanchyan
'Created date: 09/04/2007
'Last updated: 09/17/2007

Imports Microsoft.Win32
Imports System.Globalization
Imports System.Runtime.InteropServices

''' <summary>
''' Represents a time zone and provides access to all system time zones.
''' </summary>
<DebuggerDisplay("{_displayName}")> _
Public Class TimeZoneInfo : Implements IComparer(Of TimeZoneInfo)


    Private _id As String
    Private _tzi As New TimeZoneInformation
    Private _displayName As String

#Region " STRUCTURES "

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SYSTEMTIME
        Public wYear As UShort
        Public wMonth As UShort
        Public wDayOfWeek As UShort
        Public wDay As UShort
        Public wHour As UShort
        Public wMinute As UShort
        Public wSecond As UShort
        Public wMilliseconds As UShort

        ''' <summary>
        ''' Sets the member values of the time structure.
        ''' </summary>
        ''' <param name="info">A byte array that contains the information of a time.</param>
        <DebuggerHidden()> _
        Public Sub SetInfo(ByVal info() As Byte)
            If info.Length <> Marshal.SizeOf(Me) Then
                Throw New ArgumentException("Information size is incorrect", "info")
            End If
            Me.wYear = BitConverter.ToUInt16(info, 0)
            Me.wMonth = BitConverter.ToUInt16(info, 2)
            Me.wDayOfWeek = BitConverter.ToUInt16(info, 4)
            Me.wDay = BitConverter.ToUInt16(info, 6)
            Me.wHour = BitConverter.ToUInt16(info, 8)
            Me.wMinute = BitConverter.ToUInt16(info, 10)
            Me.wSecond = BitConverter.ToUInt16(info, 12)
            Me.wMilliseconds = BitConverter.ToUInt16(info, 14)
        End Sub

        ''' <summary>
        ''' Determines whether the specified System.Object 
        ''' is equal to the current System.Object.
        ''' </summary>
        ''' <param name="obj">The System.Object to compare 
        ''' with the current System.Object.</param>
        <DebuggerHidden()> _
            Public Overrides Function Equals(ByVal obj As Object) As Boolean
            If Me.GetType Is obj.GetType Then
                Dim objSt As SYSTEMTIME = DirectCast(obj, SYSTEMTIME)
                If Me.wDay <> objSt.wDay _
                OrElse Me.wDayOfWeek <> objSt.wDayOfWeek _
                OrElse Me.wHour <> objSt.wHour _
                OrElse Me.wMilliseconds <> objSt.wMilliseconds _
                OrElse Me.wMinute <> objSt.wMinute _
                OrElse Me.wMonth <> objSt.wMonth _
                OrElse Me.wSecond <> objSt.wSecond _
                OrElse Me.wYear <> objSt.wYear Then
                    Return False
                Else
                    Return True
                End If
            End If
            Return False
        End Function

    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure TimeZoneInformation
        Public bias As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public standardName As String
        Public standardDate As SYSTEMTIME
        Public standardBias As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public daylightName As String
        Public daylightDate As SYSTEMTIME
        Public daylightBias As Integer

        ''' <summary>
        ''' Sets the member values of bias, standardBias, 
        ''' daylightBias, standardDate, daylightDate of the structure.
        ''' </summary>
        ''' <param name="info">A byte array that contains the 
        ''' information of the Tzi windows registry key.</param>
        <DebuggerHidden()> _
        Public Sub SetBytes(ByVal info() As Byte)
            If info.Length <> 44 Then
                Throw New ArgumentException("Information size is incorrect", "info")
            End If
            Me.bias = BitConverter.ToInt32(info, 0)
            Me.standardBias = BitConverter.ToInt32(info, 4)
            Me.daylightBias = BitConverter.ToInt32(info, 8)
            Dim helper(15) As Byte
            Array.Copy(info, 12, helper, 0, 16)
            Me.standardDate.SetInfo(helper)
            Array.Copy(info, 28, helper, 0, 16)
            Me.daylightDate.SetInfo(helper)
        End Sub

        ''' <summary>
        ''' Determines whether the specified System.Object 
        ''' is equal to the current System.Object.
        ''' </summary>
        ''' <param name="obj">The System.Object to compare 
        ''' with the current System.Object.</param>
        <DebuggerHidden()> _
            Public Overrides Function Equals(ByVal obj As Object) As Boolean
            If Me.GetType Is obj.GetType Then
                Dim objTzi As TimeZoneInformation = DirectCast(obj, TimeZoneInformation)
                If Me.bias <> objTzi.bias _
                OrElse Me.daylightBias <> objTzi.daylightBias _
                OrElse Me.daylightName <> objTzi.daylightName _
                OrElse Me.standardBias <> objTzi.standardBias _
                OrElse Me.standardName <> objTzi.standardName _
                OrElse Not Me.daylightDate.Equals(objTzi.daylightDate) _
                OrElse Not Me.standardDate.Equals(objTzi.standardDate) Then
                    Return False
                Else
                    Return True
                End If
            End If
            Return False
        End Function

    End Structure

#End Region

#Region " API METHODS "

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, setLastError:=True)> _
    Private Shared Function SetTimeZoneInformation( _
    ByRef lpTimeZoneInformation As TimeZoneInformation) As Boolean
    End Function

#End Region

#Region " CLASS PROPERTIES "

    ''' <summary>
    ''' Gets the display name of the time zone.
    ''' </summary>
    Public ReadOnly Property DisplayName() As String
        <DebuggerHidden()> _
        Get
            Me.Refresh()
            Return Me._displayName
        End Get
    End Property

    ''' <summary>
    ''' Gets the daylight saving name of the time zone.
    ''' </summary>
    Public ReadOnly Property DaylightName() As String
        <DebuggerHidden()> _
        Get
            Me.Refresh()
            If Me.GetDaylightChanges(Me.CurrentTime.Year).Delta = TimeSpan.Zero Then
                Return Me._tzi.standardName
            Else
                Return Me._tzi.daylightName
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets the standard name of the time zone.
    ''' </summary>
    Public ReadOnly Property StandardName() As String
        <DebuggerHidden()> _
        Get
            Me.Refresh()
            Return Me._tzi.standardName
        End Get
    End Property

    ''' <summary>
    ''' Gets the current date and time of the time zone.
    ''' </summary>
    Public ReadOnly Property CurrentTime() As DateTime
        <DebuggerHidden()> _
        Get
            Return New DateTime( _
            DateTime.UtcNow.Ticks + Me.CurrentUtcOffset.Ticks, DateTimeKind.Local)
        End Get
    End Property

    ''' <summary>
    ''' Gets the current UTC (Coordinated Universal Time) offset of the time zone.
    ''' </summary>
    Public ReadOnly Property CurrentUtcOffset() As TimeSpan
        <DebuggerHidden()> _
        Get
            If Me.IsDaylightSavingTime Then
                Return New TimeSpan(0, -(Me._tzi.bias + Me._tzi.daylightBias), 0)
            Else
                Return New TimeSpan(0, -Me._tzi.bias, 0)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the current time zone for this computer system.
    ''' </summary>
    Public Shared Property CurrentTimeZone() As TimeZoneInfo
        <DebuggerHidden()> _
        Get
            Return New TimeZoneInfo(TimeZone.CurrentTimeZone.StandardName)
        End Get
        <DebuggerHidden()> _
        Set(ByVal value As TimeZoneInfo)
            value.Refresh()
            If Not TimeZoneInfo.SetTimeZoneInformation(value._tzi) Then
                'Throw a Win32Exception
                Throw New System.ComponentModel.Win32Exception()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets the standard UTC (Coordinated Universal Time) offset of the time zone.
    ''' </summary>
    Public ReadOnly Property StandardUtcOffset() As TimeSpan
        <DebuggerHidden()> _
        Get
            Me.Refresh()
            Return New TimeSpan(0, -Me._tzi.bias, 0)
        End Get
    End Property

    ''' <summary>
    ''' Gets the id of the time zone.
    ''' </summary>
    Public ReadOnly Property Id() As String
        <DebuggerHidden()> _
        Get
            Me.Refresh()
            Return Me._id
        End Get
    End Property

#End Region

#Region " CLASS CONSTRUCTORS "

    ''' <param name="standardName">A time zone standard name.</param>
    <DebuggerHidden()> _
    Public Sub New(ByVal standardName As String)
        Me.SetValues(standardName)
    End Sub

    <DebuggerHidden()> _
    Private Sub New()
    End Sub

#End Region

#Region " CLASS METHODS "

    ''' <summary>
    ''' Gets an array of all time zones on the system.
    ''' </summary>
    <DebuggerHidden()> _
    Public Shared Function GetTimeZones() As TimeZoneInfo()
        Dim tzInfos As New List(Of TimeZoneInfo)
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey( _
        "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones", False)
        If Not key Is Nothing Then
            For Each zoneName As String In key.GetSubKeyNames
                Dim tzi As TimeZoneInfo = New TimeZoneInfo
                tzi._id = zoneName
                tzi.SetValues()
                tzInfos.Add(tzi)
            Next
            TimeZoneInfo.Sort(tzInfos)
        Else
            Throw New KeyNotFoundException( _
            "Cannot find the windows registry key (Time Zone).")
        End If
        Return tzInfos.ToArray
    End Function

    ''' <summary>
    ''' Sorts the elements in a list(Of TimeZoneInfo) 
    ''' object based on standard UTC offset or display name.
    ''' </summary>
    ''' <param name="tzInfos">A time zone list to sort.</param>
    <DebuggerHidden()> _
    Public Overloads Shared Sub Sort(ByVal tzInfos As List(Of TimeZoneInfo))
        tzInfos.Sort(New TimeZoneInfo)
    End Sub

    ''' <summary>
    ''' Sorts the elements in an entire one-dimensional TimeZoneInfo 
    ''' array based on standard UTC offset or display name.
    ''' </summary>
    ''' <param name="tzInfos">A time zone array to sort.</param>
    <DebuggerHidden()> _
    Public Overloads Shared Sub Sort(ByVal tzInfos() As TimeZoneInfo)
        Array.Sort(tzInfos, New TimeZoneInfo)
    End Sub

    ''' <summary>
    ''' Gets a TimeZoneInfo.Object from standard name.
    ''' </summary>
    ''' <param name="standardName">A time zone standard name.</param>
    <DebuggerHidden()> _
    Public Shared Function FromStandardName(ByVal standardName As String) As TimeZoneInfo
        Return New TimeZoneInfo(standardName)
    End Function

    ''' <summary>
    ''' Gets a TimeZoneInfo.Object from Id.
    ''' </summary>
    ''' <param name="id">A time zone id that corresponds 
    ''' to the windows registry time zone key.</param>
    <DebuggerHidden()> _
    Public Shared Function FromId(ByVal id As String) As TimeZoneInfo
        If Not id Is Nothing Then
            If id <> String.Empty Then
                Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey( _
                "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones", False)
                If Not key Is Nothing Then
                    Dim subKey As RegistryKey = key.OpenSubKey(id, False)
                    If Not subKey Is Nothing Then
                        Dim tzi As New TimeZoneInfo
                        tzi._id = subKey.Name
                        tzi._displayName = CStr(subKey.GetValue("Display"))
                        tzi._tzi.daylightName = CStr(subKey.GetValue("Dlt"))
                        tzi._tzi.standardName = CStr(subKey.GetValue("Std"))
                        tzi._tzi.SetBytes(CType(subKey.GetValue("Tzi"), Byte()))
                        Return tzi
                    End If
                Else
                    Throw New KeyNotFoundException( _
                    "Cannot find the windows registry key (Time Zone).")
                End If
            End If
            Throw New ArgumentException("Unknown time zone.", "id")
        Else
            Throw New ArgumentNullException("id", "Value cannot be null.")
        End If
    End Function

    ''' <summary>
    ''' Returns the daylight saving time for a particular year.
    ''' </summary>
    ''' <param name="year">The year to which the daylight 
    ''' saving time period applies.</param>
    <DebuggerHidden()> _
    Public Function GetDaylightChanges( _
    ByVal year As Integer) As System.Globalization.DaylightTime
        Dim tzi As New TimeZoneInformation
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey( _
        "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones", False)
        If Not key Is Nothing Then
            Dim subKey As RegistryKey = key.OpenSubKey(Me._id, False)
            If Not subKey Is Nothing Then
                Dim subKey1 As RegistryKey = subKey.OpenSubKey("Dynamic DST", False)
                If Not subKey1 Is Nothing Then
                    If Array.IndexOf(subKey1.GetValueNames, CStr(year)) <> -1 Then
                        tzi.SetBytes(CType(subKey1.GetValue(CStr(year)), Byte()))
                    Else
                        Me.Refresh()
                        tzi = Me._tzi
                    End If
                Else
                    Me.Refresh()
                    tzi = Me._tzi
                End If
            Else
                Throw New Exception("Unknown time zone.")
            End If
        Else
            Throw New KeyNotFoundException( _
            "Cannot find the windows registry key (Time Zone).")
        End If
        Dim dStart, dEnd As DateTime
        dStart = Me.GetStartDate(tzi, year)
        dEnd = Me.GetEndDate(tzi, year)
        If dStart <> Date.MinValue AndAlso dEnd <> Date.MinValue Then
            Return New DaylightTime( _
            dStart, dEnd, New TimeSpan(0, -Me._tzi.daylightBias, 0))
        Else
            Return New DaylightTime(dStart, dEnd, New TimeSpan(0, 0, 0))
        End If
    End Function

    ''' <summary>
    ''' Returns a value indicating whether this time 
    ''' zone is within a daylight saving time period.
    ''' </summary>
    <DebuggerHidden()> _
    Public Function IsDaylightSavingTime() As Boolean
        Dim dUtcNow As DateTime = DateTime.UtcNow.AddMinutes(-(Me._tzi.bias))
        Dim sUtcNow As DateTime = DateTime.UtcNow.AddMinutes(-(Me._tzi.bias + Me._tzi.daylightBias))
        Dim dt As DaylightTime

        If Me._tzi.daylightDate.wMonth <= Me._tzi.standardDate.wMonth Then
            'Daylight saving time starts and ends in the same year
            dt = Me.GetDaylightChanges(dUtcNow.Year)
            If dt.Delta <> TimeSpan.Zero Then
                If dUtcNow >= dt.Start AndAlso sUtcNow < dt.End Then
                    Return True
                Else
                    Return False
                End If
            End If
        Else
            'Daylight saving time starts and ends in diferent years
            dt = Me.GetDaylightChanges(sUtcNow.Year)
            If dt.Delta <> TimeSpan.Zero Then
                If dUtcNow < dt.Start AndAlso sUtcNow >= dt.End Then
                    Return False
                Else
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Creates and returns a date and time object.
    ''' </summary>
    ''' <param name="wYear">The year of the date.</param>
    ''' <param name="wMonth">The month of the date.</param>
    ''' <param name="wDay">The week day in the month.</param>
    ''' <param name="wDayOfWeek">The day of the week.</param>
    ''' <param name="wHour">The hour of the date.</param>
    ''' <param name="wMinute">The minute of the date.</param>
    ''' <param name="wSecond">The seconds of the date.</param>
    ''' <param name="wMilliseconds">The milliseconds of the date.</param>
    <DebuggerHidden()> _
    Private Function CreateDate( _
    ByVal wYear As Integer, _
    ByVal wMonth As Integer, _
    ByVal wDay As Integer, _
    ByVal wDayOfWeek As Integer, _
    ByVal wHour As Integer, _
    ByVal wMinute As Integer, _
    ByVal wSecond As Integer, _
    ByVal wMilliseconds As Integer) As DateTime

        If wDay < 1 OrElse wDay > 5 Then
            Throw New ArgumentOutOfRangeException( _
            "wDat", wDay, "The value is out of acceptable range (1 to 5).")
        End If
        If wDayOfWeek < 0 OrElse wDayOfWeek > 6 Then
            Throw New ArgumentOutOfRangeException( _
            "wDayOfWeek", wDayOfWeek, "The value is out of acceptable range (0 to 6).")
        End If
        Dim daysInMonth As Integer = Date.DaysInMonth(wYear, wMonth)
        Dim fDayOfWeek As Integer = New DateTime(wYear, wMonth, 1).DayOfWeek
        Dim occurre As Integer = 1
        Dim day As Integer = 1
        If fDayOfWeek <> wDayOfWeek Then
            If wDayOfWeek = 0 Then
                day += 7 - fDayOfWeek
            Else
                If wDayOfWeek > fDayOfWeek Then
                    day += wDayOfWeek - fDayOfWeek
                ElseIf wDayOfWeek < fDayOfWeek Then
                    day = wDayOfWeek + fDayOfWeek
                End If
            End If
        End If
        While occurre < wDay AndAlso day <= daysInMonth - 7
            day += 7
            occurre += 1
        End While
        Return New DateTime( _
        wYear, wMonth, day, wHour, wMinute, wSecond, wMilliseconds, DateTimeKind.Local)
    End Function

    ''' <summary>
    ''' Gets the starting daylight saving date and time for specified thime zone.
    ''' </summary>
    <DebuggerHidden()> _
    Private Function GetStartDate(ByVal tzi As TimeZoneInformation, ByVal year As Integer) As DateTime
        With tzi.daylightDate
            If .wMonth <> 0 Then
                If .wYear = 0 Then
                    Return Me.CreateDate(year, .wMonth, _
                    .wDay, .wDayOfWeek, _
                    .wHour, .wMinute, .wSecond, .wMilliseconds)
                Else
                    Return New DateTime( _
                    .wYear, .wMonth, .wDay, _
                    .wHour, .wMinute, .wSecond, .wMilliseconds, DateTimeKind.Local)
                End If
            End If
        End With
    End Function

    ''' <summary>
    ''' Gets the end date of the daylight saving time for specified thime zone.
    ''' </summary>
    <DebuggerHidden()> _
    Private Function GetEndDate(ByVal tzi As TimeZoneInformation, ByVal year As Integer) As DateTime
        With tzi.standardDate
            If .wMonth <> 0 Then
                If .wYear = 0 Then
                    Return Me.CreateDate(year, .wMonth, _
                    .wDay, .wDayOfWeek, _
                    .wHour, .wMinute, .wSecond, .wMilliseconds)
                Else
                    Return New DateTime( _
                    .wYear, .wMonth, .wDay, _
                    .wHour, .wMinute, .wSecond, .wMilliseconds, DateTimeKind.Local)
                End If
            End If
        End With
    End Function

    ''' <summary>
    ''' Refreshes the information of the time zone object.
    ''' </summary>
    <DebuggerHidden()> _
    Public Sub Refresh()
        Me.SetValues()
    End Sub

    ''' <summary>
    ''' Sets the time zone object's information.
    ''' </summary>
    <DebuggerHidden()> _
    Private Overloads Sub SetValues()
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey( _
        "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones", False)
        If Not key Is Nothing Then
            Dim subKey As RegistryKey = key.OpenSubKey(Me._id, False)
            If Not subKey Is Nothing Then
                Me._displayName = CStr(subKey.GetValue("Display"))
                Me._tzi.daylightName = CStr(subKey.GetValue("Dlt"))
                Me._tzi.standardName = CStr(subKey.GetValue("Std"))
                Me._tzi.SetBytes(CType(subKey.GetValue("Tzi"), Byte()))
            Else
                Throw New Exception("Unknown time zone.")
            End If
        Else
            Throw New KeyNotFoundException( _
            "Cannot find the windows registry key (Time Zone).")
        End If
    End Sub

    ''' <summary>
    ''' Sets the time zone object's information.
    ''' </summary>
    ''' <param name="standardName">A time zone standard name.</param>
    <DebuggerHidden()> _
    Private Overloads Sub SetValues(ByVal standardName As String)
        If Not standardName Is Nothing Then
            Dim exist As Boolean = False
            If standardName <> String.Empty Then
                Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey( _
                "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones", False)
                If Not key Is Nothing Then
                    For Each zoneName As String In key.GetSubKeyNames
                        Dim subKey As RegistryKey = key.OpenSubKey(zoneName, False)
                        If CStr(subKey.GetValue("Std")) = standardName Then
                            Me._id = zoneName
                            Me._displayName = CStr(subKey.GetValue("Display"))
                            Me._tzi.daylightName = CStr(subKey.GetValue("Dlt"))
                            Me._tzi.standardName = CStr(subKey.GetValue("Std"))
                            Me._tzi.SetBytes(CType(subKey.GetValue("Tzi"), Byte()))
                            exist = True
                            Exit For
                        End If
                    Next
                Else
                    Throw New KeyNotFoundException( _
                    "Cannot find the windows registry key (Time Zone).")
                End If
            End If
            If Not exist Then
                Throw New ArgumentException("Unknown time zone.", "standardName")
            End If
        Else
            Throw New ArgumentNullException("id", "Value cannot be null.")
        End If
    End Sub

    ''' <summary>
    ''' Returns a System.String that represents the current TimeZoneInfo object.
    ''' </summary>
    <DebuggerHidden()> _
    Public Overrides Function ToString() As String
        Return Me.DisplayName
    End Function

    ''' <summary>
    ''' Determines whether the specified System.Object 
    ''' is equal to the current System.Object.
    ''' </summary>
    ''' <param name="obj">The System.Object to compare 
    ''' with the current System.Object.</param>
    <DebuggerHidden()> _
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If Me.GetType Is obj.GetType Then
            Dim objTzi As TimeZoneInfo = DirectCast(obj, TimeZoneInfo)
            If Me._displayName <> objTzi._displayName _
            OrElse Me._id <> objTzi._id _
            OrElse Not Me._tzi.Equals(objTzi._tzi) Then
                Return False
            Else
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Compares two specified TimeZoneInfo.Objects 
    ''' based on standard UTC offset or display name.
    ''' </summary>
    ''' <param name="x">The first TimeZoneInfo.Object.</param>
    ''' <param name="y">The second TimeZoneInfo.Object.</param>
    <DebuggerHidden()> _
    Protected Overridable Function Compare(ByVal x As TimeZoneInfo, ByVal y As TimeZoneInfo) As Integer Implements System.Collections.Generic.IComparer(Of TimeZoneInfo).Compare
        If x._tzi.bias = y._tzi.bias Then
            Return x._displayName.CompareTo(y._displayName)
        End If
        If x._tzi.bias > y._tzi.bias Then
            Return -1
        End If
        If x._tzi.bias < y._tzi.bias Then
            Return 1
        End If
    End Function

#End Region

End Class