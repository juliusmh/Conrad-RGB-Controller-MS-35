
Imports MS_RGB
Imports System.IO
Imports System.IO.Ports.SerialPort
Imports System.Threading



Module Module1
    Dim sport As New System.IO.Ports.SerialPort
    Dim mode As Integer = 0
  
    Sub Main()

        Console.WriteLine("Choose a mode to visualize")
start:
        Console.Write("COM")
        Dim com = "COM" & Console.ReadLine
        If (MS_RGB.Controls.connect(com, sport) = True) Then 'Returns true if connected
            Console.WriteLine("Connecting Succesfull")
            MS_RGB.Controls.colorchange(255, 0, 0, sport) 'Writing Bright Red to the Strip
            Thread.Sleep(1000)
            MS_RGB.Controls.colorchange(0, 255, 0, sport) 'Writing Bright Green to the Strip
            Thread.Sleep(1000)
            MS_RGB.Controls.colorchange(0, 0, 255, sport) 'Writing Bright Blue to the Strip
            Thread.Sleep(1000)
            MS_RGB.Controls.colorchange(0, 0, 0, sport) 'Writing black to the Strip
            Thread.Sleep(1000)

        Else
            Console.WriteLine("Connecting Not Succesfull")
            Console.Read()
    
        End If
        Console.Read()


    End Sub


End Module
