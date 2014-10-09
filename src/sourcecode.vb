'DLL to Control the MSRGB-35 
'By Julius HInze 
'Copyright 2013
'some edit

Imports System.IO
Imports System.Threading

Public Class Controls
    'DLL to Control the MSRGB-35 
    'By Julius HInze 
    'Copyright 2013

    'Dim serialport1 As System.IO.Ports.SerialPort
    'Klasse um den Rgb Controller MS-35 von Conrad zu steuern
    Public Shared Function connect(ByVal com As String, ByVal serialport1 As System.IO.Ports.SerialPort)

        Dim answer As Char
        serialport1.PortName = com  '<<<< Com Port
        serialport1.BaudRate = 38400 'The Baudrate (always this)
        serialport1.Parity = IO.Ports.Parity.None
        serialport1.DataBits = 8
        serialport1.StopBits = IO.Ports.StopBits.One
        serialport1.ReadTimeout = 600  '
        serialport1.WriteTimeout = -1 '
        serialport1.Handshake = IO.Ports.Handshake.None
        Try
            serialport1.Open() 'Open Up the SerialPort
        Catch ex As Exception
            Return (False)
            GoTo ending
        End Try


        clean(serialport1) 'Clean Buffers

        Dim outbyte() As Byte = {253}
        For i = 1 To 9
            serialport1.Write(outbyte, 0, 1)
            Thread.Sleep(1) 'Waiting for the device 1 ms
        Next

        clean(serialport1) 'Clean Buffers

        Dim outbyte_2() As Byte = {&HFD, &H0, &H0, &H0, &H0, &H0, &H0, &HCF, &H2C}
        'Not used because i needed the interval of 1 ms between the bytes
        Dim fdbyte() As Byte = {&HFD} 'one FD byte =D

        'Connecting sequence
        serialport1.Write(fdbyte, 0, 1) ' One FD Byte
        Thread.Sleep(1)

        For i = 1 To 6
            serialport1.Write(New Byte() {&H0}, 0, 1) ' Six HO Bytes
            Thread.Sleep(1)
        Next

        serialport1.Write(New Byte() {&HCF}, 0, 1) 'The CRC 16 LowByte
        Thread.Sleep(1)
        serialport1.Write(New Byte() {&H2C}, 0, 1) 'The CRC 16 HighByte
        Thread.Sleep(1)
        'End connecting sequence
        answer = serialport1.ReadExisting 'Read the in Buffer

        If answer = "a" Then 'If answer = e then start the native CONRAD APPLICATION and then try again!
            'the native Conrad APP resets the RGB Controller..so please do this
            Return (True)
        Else
            Return (False)
        End If
ending:
    End Function

    Public Shared Sub clean(ByVal serialport1 As System.IO.Ports.SerialPort)
        '-----Cleaning the Buffer
        serialport1.DiscardOutBuffer()
        serialport1.DiscardInBuffer()
        '-----Finish Cleaning
    End Sub


    Public Shared Function CRC16new(ByVal data() As Byte)
        Dim CRC16Lo As Byte, CRC16Hi As Byte 'CRC register
        Dim CL As Byte, CH As Byte 'Polynomial codes & HA001
        Dim SaveHi As Byte, SaveLo As Byte
        Dim i As Integer
        Dim Flag As Integer

        CRC16Lo = &H0 '&HFF
        CRC16Hi = &H0 '&HFF
        CL = &H1
        CH = &HA0

        For i = 0 To UBound(data)
            CRC16Lo = CRC16Lo Xor data(i) 'for each data and CRC register XOR

            For Flag = 0 To 7
                SaveHi = CRC16Hi
                SaveLo = CRC16Lo
                CRC16Hi = CRC16Hi \ 2 'peak shift to the right one
                CRC16Lo = CRC16Lo \ 2 'shift to the right a low

                If ((SaveHi And &H1) = &H1) Then 'If the high byte last one for a
                    CRC16Lo = CRC16Lo Or &H80 'then the low byte shifted to the right after the meeting in front of a
                End If 'Otherwise, auto-fill 0

                If ((SaveLo And &H1) = &H1) Then 'If the LSB is 1, then XOR with the polynomial codes
                    CRC16Hi = CRC16Hi Xor CH
                    CRC16Lo = CRC16Lo Xor CL
                End If
            Next Flag
        Next i

        Dim ReturnData(1) As Byte
        ReturnData(0) = CRC16Hi 'CRC high
        ReturnData(1) = CRC16Lo 'CRC low
        CRC16new = ReturnData(0) & ":" & ReturnData(1)
    End Function

    Public Shared Function colorchange(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal serialport1 As System.IO.Ports.SerialPort)
        Try
            Dim hexr, hexg, hexb As Byte
            hexr = Convert.ToByte(r)
            hexg = Convert.ToByte(g)
            hexb = Convert.ToByte(b)
            Dim crc As String
            Dim lowbyte, highbyte As Byte
            Dim splitted() As String
            Dim bytes() As Byte = {&H1, &H0, hexr, hexg, hexb, &H0, &H0}
            crc = CRC16new(bytes)
            splitted = Split(crc, ":")
            lowbyte = Convert.ToByte(splitted(0))
            highbyte = Convert.ToByte(splitted(1))
            ' 01 00 FF FF FF 00 00 F0 04 = exemple bytestream of an color change to RGB = 255

            serialport1.Write(New Byte() {&H1}, 0, 1) 'No idea why this byte
            Thread.Sleep(1)
            serialport1.Write(New Byte() {&H0}, 0, 1) 'No idea why this byte
            Thread.Sleep(1)
            serialport1.Write(New Byte() {hexr}, 0, 1) 'R
            Thread.Sleep(1)
            serialport1.Write(New Byte() {hexg}, 0, 1) 'G
            Thread.Sleep(1)
            serialport1.Write(New Byte() {hexb}, 0, 1) 'B
            Thread.Sleep(1)
            serialport1.Write(New Byte() {&H0}, 0, 1) 'No idea why this byte
            Thread.Sleep(1)
            serialport1.Write(New Byte() {&H0}, 0, 1) 'No idea why this byte
            Thread.Sleep(1)

            serialport1.Write(New Byte() {lowbyte}, 0, 1) 'CRC16 first byte
            Thread.Sleep(1)
            serialport1.Write(New Byte() {highbyte}, 0, 1) 'CRC16 second byte
            Thread.Sleep(1)
            Return (True)
        Catch ex As Exception
            Return (False)
        End Try
      
    End Function

End Class
'DLL to Control the MSRGB-35 
'By Julius HInze 
'Copyright 2013/14

