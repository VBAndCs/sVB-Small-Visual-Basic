Imports System.Globalization
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' This class provides methods to deal with date and time
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class [Date]

        ''' <summary>
        ''' Returns the current date and time as defiend by the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Date)>
        Public Shared ReadOnly Property Now As Primitive
            Get
                Return New Primitive(DateTime.Now.Ticks(), NumberType.Date)
            End Get
        End Property

        ''' <summary>
        ''' Creates a time span from given duration parts.
        ''' </summary>
        ''' <param name="days">the total days in the duration. For example, 1000 days approximatly represents 2 years and 9 months</param>
        ''' <param name="hours">the hours part of the duration</param>
        ''' <param name="minutes">the minutes part of the duration</param>
        ''' <param name="seconds">the seconds part of the duration</param>
        ''' <param name="milliseconds">the milliseconds part of the duration</param>
        ''' <returns>a new duration</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function CreateDuration(days As Primitive, hours As Primitive, minutes As Primitive, seconds As Primitive, milliseconds As Primitive) As Primitive
            Return New Primitive(New TimeSpan(days, hours, minutes, seconds, milliseconds).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creats a new date from the given text if its format is a valid date format for the given culture.
        ''' </summary>
        ''' <param name="dateText">the text that represents the date</param>
        ''' <param name="cultureName">the culture name used to format the date, like "en-us" for English United States culture, "ar-eg" for Arabic Egypt culture, and "ar-sa" for Arabic Saudi Arabia culture</param>
        ''' <returns>a new date or empty string</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function FromCulture(dateText As Primitive, cultureName As Primitive) As Primitive
            Dim d As Date
            Dim c As CultureInfo
            Try
                c = New CultureInfo(cultureName.AsString())
            Catch
                c = CultureInfo.CurrentUICulture
            End Try

            If Date.TryParse(dateText, c, DateTimeStyles.None, d) Then
                Return New Primitive(d.Ticks, NumberType.Date)
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Gets the string representaion on the given date formatted with the given culture.
        ''' </summary>
        ''' <param name="date">the input date</param>
        ''' <param name="cultureName">the culture name used to format the date, like "en-us" for English United States culture, "ar-eg" for Arabic Egypt culture, and "ar-sa" for Arabic Saudi Arabia culture</param>
        ''' <returns>a string represent the date in the given culture</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function ToCulture([date] As Primitive, cultureName As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Try
                Return d.Value.ToString(New CultureInfo(cultureName.AsString()))
            Catch
            End Try
            Return d.Value.ToString()
        End Function


        ''' <summary>
        ''' Creats a new date from the given ticks value. Note that the second contains 10 milion ticks.
        ''' </summary>
        ''' <param name="ticks">the total ticks of the date</param>
        ''' <returns>a new date</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function TicksToDate(ticks As Primitive) As Primitive
            If Not ticks.IsNumber Then Return ""
            Return New Primitive(ticks.AsDecimal, NumberType.Date)
        End Function

        ''' <summary>
        ''' Creats a new TimeSpan from the given ticks value. Note that the second contains 10 milion ticks.
        ''' </summary>
        ''' <param name="ticks">the total ticks of the date</param>
        ''' <returns>a new duration</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function TicksToDuration(ticks As Primitive) As Primitive
            If Not ticks.IsNumber Then Return ""
            Return New Primitive(ticks.AsDecimal, NumberType.TimeSpan)
        End Function


        ''' <summary>
        ''' Creats a new date from the given year, month and day values.
        ''' </summary>
        ''' <param name="year">the year number</param>
        ''' <param name="month">a number between 1 and 12 representing the month</param>
        ''' <param name="day">a number between 1 and 31 representing the day</param>
        ''' <returns>a new date</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function FromDate(year As Primitive, month As Primitive, day As Primitive) As Primitive
            Try
                Dim d As New Date(year, month, day)
                Return New Primitive(d.Ticks, NumberType.Date)
            Catch
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' Creats a new date from the given time values.
        ''' </summary>
        ''' <param name="hour">a number between 0 and 23 representing the hour</param>
        ''' <param name="minute">a number between 0 and 59 representing the minute</param>
        ''' <param name="second">a number between 0 and 59 representing the second</param>
        ''' <param name="millisecond">a number between 0 and 999 representing the millisecond</param>
        ''' <returns>a new date</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function FromTime(hour As Primitive, minute As Primitive, second As Primitive, millisecond As Primitive) As Primitive
            Try
                Dim d As New Date(1, 1, 1, hour, minute, second, millisecond)
                Return New Primitive(d.Ticks, NumberType.Date)
            Catch
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' Creats a new date from the given date and time values.
        ''' </summary>
        ''' <param name="year">the year number</param>
        ''' <param name="month">a number between 1 and 12 representing the month</param>
        ''' <param name="day">a number between 1 and 31 representing the day</param>
        ''' <param name="hour">a number between 0 and 23 representing the hour</param>
        ''' <param name="minute">a number between 0 and 59 representing the minute</param>
        ''' <param name="second">a number between 0 and 59 representing the second</param>
        ''' <param name="millisecond">a number between 0 and 999 representing the millisecond</param>
        ''' <returns>a new date</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function FromDateTime(year As Primitive, month As Primitive, day As Primitive, hour As Primitive, minute As Primitive, second As Primitive, millisecond As Primitive) As Primitive
            Try
                Dim d As New Date(year, month, day, hour, minute, second, millisecond)
                Return New Primitive(d.Ticks, NumberType.Date)
            Catch
            End Try
            Return ""
        End Function


        ''' <summary>
        ''' Gets the full time part (including seconds and milliseconds) of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the long time</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetLongTime([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToLongTimeString())
        End Function

        ''' <summary>
        ''' Gets the short time part of the given date. The time will incude hours and minutes and AM or PM, but not seconds and milliseconds.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the short time</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetShortTime([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToShortTimeString())
        End Function

        ''' <summary>
        ''' Gets the long form of the given date. The long date contains the month name instead of its number.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the long date</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetLongDate([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToLongDateString())
        End Function

        ''' <summary>
        ''' Gets the short form of the given date, like 1/1/2020
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the short date</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetShortDate([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToShortDateString())
        End Function

        ''' <summary>
        ''' Gets the short date and time representaion
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the full date and time</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetLongDateAndTime([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToString())
        End Function

        ''' <summary>
        ''' Gets the long date and time representaion
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a string representing the long date and time</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetDateAndTime([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToLongDateString() & " " & d.Value.ToLongTimeString())
        End Function

        ''' <summary>
        ''' Gets the year or the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number representing the year</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetYear([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Year)
        End Function

        ''' <summary>
        ''' Gets the month of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 1 and 12 that represents the month</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetMonth([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Month)
        End Function

        ''' <summary>
        ''' Gets the English name of the month of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the name of the month in English</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetEnglishMonthName([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToString("MMMM", CultureInfo.InvariantCulture))
        End Function

        ''' <summary>
        ''' Gets the local name of the month of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the name of the month in the local language defiend on the user system</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetMonthName([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToString("MMMM"))
        End Function

        ''' <summary>
        ''' Gets the day of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a nimber between 1 and 31 that represents the day</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDay([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Day)
        End Function

        ''' <summary>
        ''' Gets the day number in the year of the given date. 
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 1 and 366 that represnts the date in the week</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDayOfYear([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.DayOfYear)
        End Function


        ''' <summary>
        ''' Gets the day of the week of the given date. Note that Sunday is the first day of the week. 
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 1 and 7 that represnts the date in the week</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDayOfWeek([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(1 + d.Value.DayOfWeek)
        End Function


        ''' <summary>
        ''' Gets the English name of the week day of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the name of the week day in English</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetEnglishDayName([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.DayOfWeek.ToString())
        End Function

        ''' <summary>
        ''' Gets the local name of the week day of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the name of the week day in the local language defiend on the user system</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetDayName([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.ToString("dddd"))
        End Function

        ''' <summary>
        ''' Gets the hour of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 0 and 23 that represents the hour</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetHour([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Hour)
        End Function


        ''' <summary>
        ''' Gets the minute of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 0 and 59 that represents the minute</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetMinute([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Minute)
        End Function


        ''' <summary>
        ''' Gets the second of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 0 and 59 that represents the second</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetSecond([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Second)
        End Function

        ''' <summary>
        ''' Gets the millisecond of the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>a number between 0 and 999 that represents the millisecond</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetMillisecond([date] As Primitive) As Primitive
            Dim d = [date].AsDate()
            If d Is Nothing Then Return ""
            Return New Primitive(d.Value.Millisecond)
        End Function

        ''' <summary>
        ''' Gets the total ticks in the given date. Note that one second contains 10 million ticks!
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the total ticks</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTicks([date] As Primitive) As Primitive
            If [date].IsTimeSpan Then
                Dim ts = [date].AsTimeSpan()
                If ts Is Nothing Then Return ""
                Return New Primitive(ts.Value.Ticks)
            Else
                Dim d = [date].AsDate()
                If d Is Nothing Then
                    Dim ts = [date].AsTimeSpan()
                    If ts Is Nothing Then Return ""
                    Return New Primitive(ts.Value.Ticks)
                End If
                Return New Primitive(d.Value.Ticks)
            End If
        End Function

        ''' <summary>
        ''' Calculates the difference in milliseconds betwwen the start of the year 1900 and the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <returns>the number of milliseconds that have elapsed since 1900 until the given date</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetElapsedMilliseconds([date] As Primitive) As Primitive
            If [date].IsTimeSpan Then
                Return [date].AsDecimal \ 10000
            End If

            Dim d = [date].AsDate()
            If d Is Nothing Then
                Dim ts = [date].AsTimeSpan()
                If ts Is Nothing Then Return 0
                Return ts.Value.TotalMilliseconds
            End If

            Return (d.Value - New DateTime(1900, 1, 1)).TotalMilliseconds
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the year set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the new year</param>
        ''' <returns>a new date with the new year value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeYear([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(value, d1.Month, d1.Day, d1.Hour, d1.Minute, d1.Second, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the month set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 1 and 12 that represents the month</param>
        ''' <returns>a new date with the new month value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeMonth([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, value, d1.Day, d1.Hour, d1.Minute, d1.Second, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the day set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 1 and 31 that represents the day</param>
        ''' <returns>a new date with the new day value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeDay([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, d1.Month, value, d1.Hour, d1.Minute, d1.Second, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the hour set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 0 and 23 that represents the new hour</param>
        ''' <returns>a new date with the new hour value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeHour([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, d1.Month, d1.Day, value, d1.Minute, d1.Second, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the minute set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 0 and 59 that represents the new minute</param>
        ''' <returns>a new date with the new minute value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeMinute([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, d1.Month, d1.Day, d1.Hour, value, d1.Second, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the second set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 0 and 59 that represents the new second</param>
        ''' <returns>a new date with the new second value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeSecond([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, d1.Month, d1.Day, d1.Hour, d1.Minute, value, d1.Millisecond)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date based on the given date with the millisecond set to the given value.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">a number between 0 and 999 that represents the new millisecond</param>
        ''' <returns>a new date with the new millisecond value. The input date will not change</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ChangeMillisecond([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]

            Dim d = [date].AsDate()
            If d Is Nothing Then Return [date]

            Try
                Dim d1 = d.Value
                Dim d2 As New Date(d1.Year, d1.Month, d1.Day, d1.Hour, d1.Minute, d1.Second, value)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function


        ''' <summary>
        ''' Creates a new date by adding the given years to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of years you want to add. If the given date is a duration, the total days of the years will be calculated assuming that a year = 365.25 days.</param>
        ''' <returns>a new date with the added years. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddYears([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            If [date].IsTimeSpan Then
                Return [date] + New Primitive(TimeSpan.FromDays(value * 365.25).Ticks, NumberType.TimeSpan)
            End If

            Dim d = [date].AsDate()
            If d Is Nothing Then
                Dim ts = [date].AsTimeSpan()
                If ts Is Nothing Then Return [date]
                Return New Primitive(ts.Value.Ticks + TimeSpan.FromDays(value * 365.25).Ticks, NumberType.TimeSpan)
            End If

            Try
                Dim d2 = d.Value.AddYears(value)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given months to the given date. 
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of months you want to add. If the given date is a duration, the total days of the months will be calculated assuming that a month = 30.44 days in avarage.</param>
        ''' <returns>a new date with the added months. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddMonths([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            If [date].IsTimeSpan Then
                Return [date] + New Primitive(TimeSpan.FromDays(value * 365.25 / 12).Ticks, NumberType.TimeSpan)
            End If

            Dim d = [date].AsDate()
            If d Is Nothing Then
                Dim ts = [date].AsTimeSpan()
                If ts Is Nothing Then Return [date]
                Return New Primitive(ts.Value.Ticks + TimeSpan.FromDays(value * 365.25 / 12).Ticks, NumberType.TimeSpan)
            End If

            Try
                Dim d2 = d.Value.AddMonths(value)
                Return New Primitive(d2.Ticks, NumberType.Date)
            Catch
            End Try

            Return [date]
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given days to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of days you want to add</param>
        ''' <returns>a new date with the added days. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddDays([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Return [date] + New Primitive(TimeSpan.FromDays(value).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given hours to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of hours you want to add</param>
        ''' <returns>a new date with the added hours. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddHours([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Return [date] + New Primitive(TimeSpan.FromHours(value).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given minutes to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of minutes you want to add</param>
        ''' <returns>a new date with the added minutes. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddMinutes([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Return [date] + New Primitive(TimeSpan.FromMinutes(value).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given seconds to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of seconds you want to add</param>
        ''' <returns>a new date with the added seconds. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddSeconds([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Return [date] + New Primitive(TimeSpan.FromSeconds(value).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given milliseconds to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="value">the number of milliseconds to add</param>
        ''' <returns>a new date with the added milliseconds. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function AddMilliseconds([date] As Primitive, value As Primitive) As Primitive
            If Not value.IsNumber Then Return [date]
            Return [date] + New Primitive(TimeSpan.FromMilliseconds(value).Ticks, NumberType.TimeSpan)
        End Function

        ''' <summary>
        ''' Creates a new date by adding the given duration to the given date.
        ''' </summary>
        ''' <param name="date">the input date \ time</param>
        ''' <param name="duration">a date representing the duration you want to add</param>
        ''' <returns>a new date with the added duration. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function Add([date] As Primitive, duration As Primitive) As Primitive
            Return [date] + duration
        End Function

        ''' <summary>
        ''' Creates a new date by subtracting the given duration from the given date.
        ''' </summary>
        ''' <param name="date1">the input date \ time</param>
        ''' <param name="date2">the date\time you want to subtract</param>
        ''' <returns>a duration representing the difference between the two dates. The input date will not Add</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function Subtract(date1 As Primitive, date2 As Primitive) As Primitive
            Return date1 - date2
        End Function

        Private Shared Function GetTimeSpan(duration As Primitive) As TimeSpan?
            If duration.IsNumber Then
                Return New TimeSpan(duration.AsDecimal())

            Else ' Parse the string value
                Dim ts = duration.AsTimeSpan()
                If ts Is Nothing Then
                    Dim d = duration.AsDate()
                    If d Is Nothing Then Return Nothing
                    Return New TimeSpan(d.Value.Ticks)
                End If
            End If
        End Function

        ''' <summary>
        ''' Gets the approximate total years in the given duration, assuming that each year contains 365.25 days.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the approximate total years</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalYears(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalDays / 365.25)
        End Function


        ''' <summary>
        ''' Gets the approximate total months in the given duration, assuming that theere is 12 months in the year, and each year contains 365.25 days.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the approximate total months</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalMonths(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalDays * 12 / 365.25)
        End Function

        ''' <summary>
        ''' Gets the total days in the given duration.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the total days</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalDays(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalDays)
        End Function


        ''' <summary>
        ''' Gets the total hours in the given duration.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the total hours</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalHours(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalHours)
        End Function

        ''' <summary>
        ''' Gets the total minutes in the given duration.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the total minutes</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalMinutes(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalMinutes)
        End Function


        ''' <summary>
        ''' Gets the total seconds in the given duration.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the total seconds</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalSeconds(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalSeconds)
        End Function


        ''' <summary>
        ''' Gets the total milliseconds in the given duration.
        ''' </summary>
        ''' <param name="duration">the input duration</param>
        ''' <returns>the total milliseconds</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTotalMilliseconds(duration As Primitive) As Primitive
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.TotalMilliseconds)
        End Function

        ''' <summary>
        ''' Get the integral days part of the given TimeSpan.
        ''' </summary>
        ''' <param name="duration">the input timespan</param>
        ''' <returns>an integer number representing the days part of the given duration</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDurationDays(duration As Primitive)
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.Days)
        End Function

        ''' <summary>
        ''' Get the integral hours part of the gdiven TimeSpan.
        ''' </summary>
        ''' <param name="duration">the input timespan</param>
        ''' <returns>an integer number representing the hours part of the given duration</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDurationHours(duration As Primitive)
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.Hours)
        End Function

        ''' <summary>
        ''' Get the integral minutes part of the given TimeSpan.
        ''' </summary>
        ''' <param name="duration">the input timespan</param>
        ''' <returns>an integer number representing the minutes part of the given duration</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDurationMinutes(duration As Primitive)
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.Minutes)
        End Function

        ''' <summary>
        ''' Get the integral seconds part of the given TimeSpan.
        ''' </summary>
        ''' <param name="duration">the input timespan</param>
        ''' <returns>an integer number representing the seconds part of the given duration</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDurationSeconds(duration As Primitive)
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.Seconds)
        End Function

        ''' <summary>
        ''' Get the integral milliseconds part of the given TimeSpan.
        ''' </summary>
        ''' <param name="duration">the input timespan</param>
        ''' <returns>an integer number representing the milliseconds part of the given duration</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetDurationMilliseconds(duration As Primitive)
            Dim ts = GetTimeSpan(duration)
            If ts Is Nothing Then Return 0
            Return New Primitive(ts.Value.Milliseconds)
        End Function
    End Class
End Namespace
