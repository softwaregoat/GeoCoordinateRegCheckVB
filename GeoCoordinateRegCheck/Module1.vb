Imports System.Text
Imports System.Text.RegularExpressions
Imports CoordinateSharp

Module Module1
    Sub Main()
        Dim c As Coordinate
        c = New Coordinate()
        c.GeoDate = DateTime.Now
        Coordinate.TryParse("N14 15 23 E36 55 44", c)
        Console.WriteLine("Parsed String: " & c.Display)
        Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
        Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)

        Dim str As String
        While True
            Console.WriteLine("*********Starting*********")
            Console.WriteLine("Input String: ")
            str = Console.ReadLine()
            Dim rgx As Regex
            rgx = New Regex("[^a-zA-Z0-9 +.,-]")
            str = rgx.Replace(str, "")
            str = str.ToUpper()
            str = str.Trim()
            Console.WriteLine("Filtered String: " & str)
            detectCoordenatesMGRS(str)
            detectCoordenatesDegrees(str)
            detectCoordenatesDecimalDegrees(str)
            detectCoordenatesDegreeDecimalMinute(str)
            detectCoordenatesDegreeMinuteSecond(str)
        End While

        Console.ReadKey()
    End Sub
    Public Sub detectCoordenatesDegreeMinuteSecond(s As String)
        Console.WriteLine("    *****detectCoordenatesDegreeMinuteSecond*********")
        Dim pat As New StringBuilder
        'N 37º 53' 27.6" W 74º 52' 19.2" [\d.]+
        pat.Append("[NS](\s*)([-+]?)([\d.]+)(\s*)[\d.]+(\s*)[\d.]+(\s*)[EW](\s*)([-+]?)[\d.]+(\s*)([\d+]{1,2})(\s*)(([\d.]+))")
        '==============Degree Minute Second
        'N37 53' 27.6", W74 52' 19.2"
        pat.Append("|")
        pat.Append("[NS]([-+]?)([\d.]+)(\s*)([\d.]+)(\s*)[\d.]+(,)(\s*)[EW](\s*)([-+]?)([\d.]+)(\s*)([\d.]+)(\s*)(([\d.]+))")
        'N37 53 27.6" W74 52 19.2"
        pat.Append("|")
        pat.Append("[NS]([-+]?)[\d.]+(\s*)[\d.]+(\s*)[\d.]+(\s*)[EW](\s*)([-+]?)[\d.]+(\s*)([\d.]+)(\s*)(([\d.]+))")
        '37 53 27.6 N 74 52 19.2 W
        pat.Append("|")
        pat.Append("([-+]?)([\d.]+)(\s*)([\d.]+)(\s*)([\d.]+)(((\.)(\d+)(\s*)[NS]))(\s*)([-+]?)[\d.]+(\s*)[\d.]+(\s*)[\d.]+(((\.)(\d+)(\s*)[EW]))")
        '37-53-27.6 N 74-52-19.2 W
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(-)[\d.]+(-)[\d.]+(((\.)(\d+)(\s*)[NS]))(\s*)([-+]?)[\d.]+(-)[\d.]+(-)[\d.]+(((\.)(\d+)(\s*)[EW]))")
        '37-53 27.6 N 74-52 19.2 W
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(-)[\d.]+(\s*)[\d.]+(((\.)(\d+)(\s*)[NS]))(\s*)([-+]?)[\d.]+(-)[\d.]+(\s*)[\d.]+(((\.)(\d+)(\s*)[EW]))")

        Dim pattern As String = pat.ToString
        Dim m As Match = Regex.Match(s, pattern, RegexOptions.Compiled)
        While m.Success
            If m.Value.Length > 0 Then
                Try
                    Console.WriteLine("Geo string: " & m.Value)
                    Dim c As Coordinate
                    c = New Coordinate()
                    c.GeoDate = DateTime.Now
                    Coordinate.TryParse(m.Value, c)
                    Console.WriteLine("Parsed String: " & c.Display)
                    Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
                    Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)
                Catch ex As Exception
                    Console.WriteLine("Parse error: " & ex.ToString)
                End Try
            Else
                Console.WriteLine("Can not find")
                Console.WriteLine(DBNull.Value)
                Console.WriteLine(DBNull.Value)
            End If
            m = m.NextMatch
        End While

    End Sub
    Public Sub detectCoordenatesDegreeDecimalMinute(s As String)
        Console.WriteLine("    *****detectCoordenatesDegreeDecimalMinute*********")
        Dim pat As New StringBuilder
        '37º 53.46'N 74º 52.32'W
        pat.Append("([-+]?)[\d.]+(\s*)[\d.]+(((\.)(\d+)[NS]))(\s*)([-+]?)[\d.]+(\s*)[\d.]+(((\.)(\d+)[EW]))")
        '==============Degree Decimal Minute
        '37º 53.46' N, 74º 52.32' W
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(\s*)[\d.]+(((\.)(\d+)(\s*)[NS](,)))(\s*)([-+]?)[\d.]+(\s*)[\d.]+(((\.)(\d+)(\s*)[EW]))")
        '37-53.46N 74-52.32W
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(-)[\d.]+(((\.)(\d+)[NS]))(\s*)([-+]?)[\d.]+(-)[\d.]+(((\.)(\d+)(\s*)[EW]))")

        Dim pattern As String = pat.ToString
        Dim m As Match = Regex.Match(s, pattern, RegexOptions.Compiled)
        While m.Success
            If m.Value.Length > 0 Then
                Try
                    Console.WriteLine("Geo string: " & m.Value)
                    Dim c As Coordinate
                    c = New Coordinate()
                    c.GeoDate = DateTime.Now
                    Coordinate.TryParse(m.Value, c)
                    Console.WriteLine("Parsed String: " & c.Display)
                    Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
                    Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)
                Catch ex As Exception
                    Console.WriteLine("Parse error: " & ex.ToString)
                End Try
            Else
                Console.WriteLine("Can not find")
                Console.WriteLine(DBNull.Value)
                Console.WriteLine(DBNull.Value)
            End If
            m = m.NextMatch
        End While

    End Sub
    Public Sub detectCoordenatesDecimalDegrees(s As String)
        Console.WriteLine("    *****detectCoordenatesDecimalDegrees*********")
        Dim pat As New StringBuilder
        '37.891ºN 74.872ºW
        pat.Append("([-+]?)[\d.]+(((\.)(\d+)[NS]))(\s*)([-+]?)[\d.]+(((\.)(\d+)[EW]))")
        '==============Degree Decimal Minute
        'N 37º 53.46' W 74º 52.32'
        pat.Append("|")
        pat.Append("[NS]([-+ ]?)[\d.]+(\s*)[\d.]+(\s*)[EW](\s*)([-+]?)[\d.]+(\s*)(([\d.]+))")

        Dim pattern As String = pat.ToString
        Dim m As Match = Regex.Match(s, pattern, RegexOptions.Compiled)
        While m.Success
            If m.Value.Length > 0 Then
                Try
                    Console.WriteLine("Geo string: " & m.Value)
                    Dim c As Coordinate
                    c = New Coordinate()
                    c.GeoDate = DateTime.Now
                    Coordinate.TryParse(m.Value, c)
                    Console.WriteLine("Parsed String: " & c.Display)
                    Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
                    Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)
                Catch ex As Exception
                    Console.WriteLine("Parse error: " & ex.ToString)
                End Try
            Else
                Console.WriteLine("Can not find")
                Console.WriteLine(DBNull.Value)
                Console.WriteLine(DBNull.Value)
            End If
            m = m.NextMatch
        End While

    End Sub
    Public Sub detectCoordenatesDegrees(s As String)
        Console.WriteLine("    *****detectCoordenatesSignedDegrees*********")
        Dim pat As New StringBuilder
        '==============Signed Degrees
        '-51.498134, -0.201755 OK
        pat.Append("([-+]?)[\d.]+(((\.)(\d+)(,)))(\s*)(([-+]?)([\d.]+)((\.)(\d+))?)")
        '37.891 -74.872 OK
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(\s*)(([-+]?)([\d]{1,3})((\.)(\d+))?)")
        '37.891º -74.872º OK
        pat.Append("|")
        pat.Append("([-+]?)[\d.]+(((\.)(\d+)([^\u0000-\u007F])))(\s*)([-+]?)[\d.]+(((\.)(\d+)([^\u0000-\u007F])))")
        '==============Decimal Degrees
        'N 37.891, W 74.872, OK
        pat.Append("|")
        pat.Append("[NS](\s*)([-+]?)[\d.]+(((\.)(\d+)(,)))(\s*)([EW](\s*)([-+]?)([\d]{1,3})((\.)(\d+))?)")
        'N 37.891 W 74.872, OK
        pat.Append("|")
        pat.Append("[NS](\s*)([-+]?)[\d.]+(\s*)([EW](\s*)([-+]?)([\d.]+)((\.)(\d+))?)")
        'N 37.891º W 74.872º, OK
        pat.Append("|")
        pat.Append("[NS](\s*)([-+]?)[\d.]+(((\.)(\d+)([^\u0000-\u007F])))(\s*)[EW](\s*)([-+]?)[\d.]+(((\.)(\d+)([^\u0000-\u007F])))")

        Dim pattern As String = pat.ToString
        Dim m As Match = Regex.Match(s, pattern, RegexOptions.Compiled)
        While m.Success
            If m.Value.Length > 0 Then
                Try
                    Console.WriteLine("Geo string: " & m.Value)
                    Dim c As Coordinate
                    c = New Coordinate()
                    c.GeoDate = DateTime.Now
                    Coordinate.TryParse(m.Value, c)
                    Console.WriteLine("Parsed String: " & c.Display)
                    Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
                    Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)
                Catch ex As Exception
                    Console.WriteLine("Parse error: " & ex.ToString)
                End Try
            Else
                Console.WriteLine("Can not find")
                Console.WriteLine(DBNull.Value)
                Console.WriteLine(DBNull.Value)
            End If
            m = m.NextMatch
        End While

    End Sub
    Public Sub detectCoordenatesMGRS(s As String)
        Console.WriteLine("    *****detectCoordenatesMGRS*********")
        Dim pat As New StringBuilder
        pat.Append("([6][0]|[1-5][0-9]|[0]*[1-9])\s*[A-HJ-NP-Xc-hj-np-z]\s*[A-HJ-NP-Za-hj-np-z]") '[^\d]\s*
        pat.Append("{2}\s*([\d]{8}\s+[\d]{8}|[\d]{7}\s+[\d]{7}|[\d]{6}\s+[\d]{6}|[\d]{5}\s+[\d]{5}|")
        pat.Append("[\d]{4}\s+[\d]{4}|[\d]{3}\s+[\d]{3}|[\d]{2}\s+[\d]{2}|[\d]{1}\s+[\d]{1}|")
        pat.Append("\d{16}|\d{14}|\d{12}|\d{10}|\d{8}|\d{6}|\d{4}|\d{2}|)") '\s*[^\d]
        Dim pattern As String = pat.ToString
        Dim m As Match = Regex.Match(s, pattern, RegexOptions.IgnoreCase)
        While m.Success
            If m.Value.Length > 0 Then
                Try
                    Console.WriteLine("Geo string: " & m.Value)
                    Dim c As Coordinate
                    c = New Coordinate()
                    c.GeoDate = DateTime.Now
                    Coordinate.TryParse(m.Value, c)
                    Console.WriteLine("Parsed String: " & c.Display)
                    Console.WriteLine("Parsed Latitude: " & c.Latitude.ToDouble)
                    Console.WriteLine("Parsed Longitude: " & c.Longitude.ToDouble)
                Catch ex As Exception
                    Console.WriteLine("Parse error: " & ex.ToString)
                End Try
            Else
                Console.WriteLine("Can not find")
                Console.WriteLine(DBNull.Value)
                Console.WriteLine(DBNull.Value)
            End If
            m = m.NextMatch
        End While

    End Sub
End Module
