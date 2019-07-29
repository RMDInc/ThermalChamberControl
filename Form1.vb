Public Class Form1

    Dim comm As New IO.Ports.SerialPort
    Dim ports As New IO.Ports.SerialPort

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'This handles all the actions taken before Form_1 loads

        Dim portnamesarray() As String = IO.Ports.SerialPort.GetPortNames

        'For i = 0 To (portnamesarray.Length - 1)
        'can use this loop to display a list of port names
        'then use another subroutine to let users select which port they would like
        'That's probably easiest for now because it's already written
        'otherwise we don't need this...just set the port = COM1
        'Next

        Dim eurotherm_port As String = "COM1"
        If (ports.IsOpen) Then
            ports.Close()
        End If

        With ports
            .PortName = eurotherm_port
            .BaudRate = 9600
            .DataBits = 8
            .Parity = IO.Ports.Parity.None
            .StopBits = IO.Ports.StopBits.One
            .ReadTimeout = 100
            .WriteTimeout = 100
        End With

        ports.Open()
        ports.DiscardInBuffer()
        ports.DiscardOutBuffer()

        updatetextbox.Text = "The port has been set."

        File_write_timer.Interval = 5000    'Tells the system to write every 5 s (5000 ms)

    End Sub

    Private Sub runbutton_Click(sender As Object, e As EventArgs) Handles runbutton.Click
        'This handles what happens when the user clicks the "Run button" in the bottom right corner of the application

        Dim j, t0, t1, rate, hovert, cycle_number, pre As Integer
        Dim holder As Double
        Dim section_list_array() As String
        Dim textboxarray() As TextBox

        File_write_timer.Start()
        currentTemp.Text = read_eurotemp()

        textboxarray = {TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12, TextBox13, TextBox14,
            TextBox15, TextBox16, TextBox17, TextBox18, TextBox19, TextBox20, TextBox21, TextBox22, TextBox23, TextBox24, TextBox25, TextBox26, TextBox27,
            TextBox28, TextBox29, TextBox30, TextBox31, TextBox32, TextBox33, TextBox34, TextBox35, TextBox36, TextBox37}

        'Want to run a set and check that the cycle starts at T0
        'This will eventually change and I will put in a separate place to put the starting temp
        'I think this should function very close to the 'Manual set temp' button from the previous
        'iteration of this program...we'll see what users/Erik want.

        section_list_array = section_list()

        'Get the selections from the listboxes
        pre = precheckTemps()
        If pre = 0 Then
            updatetextbox.Text = "Please check T0, T1 values for accidental T0>T1 or T1>T0." & ControlChars.NewLine & "Also check that Rate and Hover time are above 0."
            Application.DoEvents()
            Return
        End If

        set_eurotemp(TextBox2.Text)
        While True
            updatetextbox.Text = "Moving to initial T0"
            currentTemp.Text = read_eurotemp()
            holder = read_eurotemp()
            If ((holder + 0.2 > TextBox2.Text) And (holder - 0.2 < TextBox2.Text)) Then
                updatetextbox.Text = "Reached initial T0" & ControlChars.NewLine & "Waiting for thermal chamber to settle."
                Application.DoEvents()
                hover(1)
                Exit While
            End If
            Application.DoEvents()
        End While

        ''''''''''''''''' This is the main loop where the cycle is performed '''''''''''''''''''''''''''''''''
        While True
            For j = 0 To section_list_array.Length - 1  'For loop goes through all sections
                updatetextbox.Text = section_list_array(j)
                Application.DoEvents()

                'Each time we increment, we need to call from different text boxes
                'To find a textbox, we start with the first box, but subtract 1 to account for the first element in the array
                t0 = 2 + (4 * j) - 1
                t1 = 3 + (4 * j) - 1
                rate = 4 + (4 * j) - 1
                hovert = 5 + (4 * j) - 1

                'Start the next selection
                If section_list_array(j) = "Ramp Up" Then
                    rampup(Convert.ToDouble(textboxarray(t0).Text), Convert.ToDouble(textboxarray(t1).Text), Convert.ToDouble(textboxarray(rate).Text))
                ElseIf section_list_array(j) = "Ramp Down" Then
                    rampdown(Convert.ToDouble(textboxarray(t0).Text), Convert.ToDouble(textboxarray(t1).Text), Convert.ToDouble(textboxarray(rate).Text))
                ElseIf section_list_array(j) = "Hover" Then
                    hover(Convert.ToDouble(textboxarray(hovert).Text))
                End If

            Next
            cycle_number += 1
            updatetextbox.Text = "Cycles performed: " & Convert.ToString(cycle_number)

            If cycle_number = cycles.Text Then  'Check to see how many cycles are left, if any
                File_write_timer.Stop()
                Exit While
            End If
        End While

    End Sub

    Function section_list() As String()
        'Goes through each list box and reads the selection into an array
        'Returns an array with entries equal to the selections made by the user

        Dim i As Integer = 8            'Start at 8 and decrement through the loop, 8 because we have 9 textboxes, including 0
        Dim todolist(8) As String   'This should have as many elements as we have listboxes


        For Each lb In Me.Controls.OfType(Of ListBox)() 'Loop over each ListBox and get the selection
            'Instead of going in ascending numerical order...vb goes in descending order...from 9->1...for no good reason
            'So, we decrement through this loop so we can read the selections from the listboxes into the correct place
            'in our array.

            If lb.Text = "" Then                    'If the box doesn't have a selection, skip it
                i = i - 1
                Continue For
            ElseIf lb.Text = " " Then               'Allows for skipping 'Null' selection boxes
                i = i - 1
                Continue For
            Else
                todolist(i) = lb.Text               'Otherwise add the text to our array
            End If

            i = i - 1
            Application.DoEvents()
        Next

        Return todolist

    End Function

    Function read_eurotemp() As Double
        'Reads the temperature from the controller; details located in original file

        Dim cmd(20) As Byte 'bytes to send to conttroller
        Dim dmc(2) As Byte  'scratch buffer
        Dim k As Integer 'loop var
        Dim crcval As Long  'function return
        Dim setpoint As Double
        Dim processvalue As Double

        cmd(0) = 1         '' Address
        cmd(1) = 3         '' Read words
        cmd(2) = 0         '' MSByte of address.
        cmd(3) = 1         '' LSbyte of address.
        cmd(4) = 0         '' MSByte of word count.
        cmd(5) = 2         '' LSByte of word count.

        crcval = CRC(cmd, 6) 'calculate checksum

        Try
            ports.Write(cmd, 0, 8)

            For k = 0 To 8
                ports.Read(cmd, k, 1)
            Next

            If CRC(cmd, 8) <> 0 Then
                MsgBox("CRC checksum FAILED")
            End If

        Catch ex As Exception
            updatetextbox.Text = "Timeout during read_temp attempt."
        End Try

        dmc(0) = cmd(4)
        dmc(1) = cmd(3)
        processvalue = CDbl(BitConverter.ToInt16(dmc, 0)) / 10.0
        dmc(0) = cmd(6)
        dmc(1) = cmd(5)
        setpoint = CDbl(BitConverter.ToInt16(dmc, 0)) / 10.0

        'Proc_val.Text = processvalue
        Return processvalue

    End Function

    Function set_eurotemp(ByVal dval As Integer)
        'Writes the temperature from the controller; details located in original file

        Dim cmd(20) As Byte 'bytes to send to conttroller
        Dim dmc(2) As Byte  'scratch buffer
        Dim k As Integer 'loop var
        Dim crcval As Long  'function return
        Dim setpoint As Double
        Dim k1 As Short

        tbsetPoint.Text = Convert.ToString(dval)

        cmd(0) = 1      ' Address inside controller to write
        cmd(1) = 6      ' Write bytes to controller
        cmd(2) = 0      ' Write 1 word
        cmd(3) = 2      ' Address of where to write (MSByte)
        setpoint = 10.0 * Val(dval)
        k1 = setpoint
        dmc = BitConverter.GetBytes(k1)
        cmd(4) = dmc(1)
        cmd(5) = dmc(0)
        crcval = CRC(cmd, 6)

        Try
            ports.Write(cmd, 0, 8)

            System.Threading.Thread.Sleep(1000)
            For k = 0 To 7
                ports.Read(cmd, k, 1)
            Next
        Catch ex As Exception
            updatetextbox.Text = "Timeout during set_temp attempt."
        End Try

        Return 1
    End Function

    Function CRC(ByVal cmd As Byte(), ByVal len As Integer) As Long
        '' CRC runs cyclic redundancy check algorithm on input message$
        '' Returns value of 16 bit CRC after completion and
        '' always adds 2 CRC bytes to message.
        '' returns 0 if incoming message has correct CRC
        ''
        '' Must use double word for CRC and decimal constants
        ''
        Dim crc16 As UShort
        Dim c As Integer
        Dim bit As Integer
        Dim crch As Byte
        Dim crcl As Byte

        crc16 = 65535
        For c = 0 To len - 1
            crc16 = crc16 Xor cmd(c)
            For bit = 1 To 8
                If crc16 Mod 2 Then
                    crc16 = (crc16 \ 2) Xor 40961
                Else
                    crc16 = crc16 \ 2
                End If
            Next bit
        Next c
        crch = crc16 \ 256
        crcl = crc16 Mod 256
        cmd(len) = crcl
        cmd(len + 1) = crch

    End Function

    Private Sub File_wrtie_timer_Tick(sender As Object, e As EventArgs) Handles File_write_timer.Tick
        'This controls how often the program writes to a file

        Dim holder As Double
        holder = read_eurotemp()
        writefile(Convert.ToString(holder))

    End Sub

    Public Sub writefile(ByVal temperature As String)
        'This takes in a temperature and prints it to an excel file along with the date and time 

        Dim filename As String
        Try
            filename = FolderBrowserDialog1.SelectedPath & "\" & logfiletextbox.Text & ".csv"        'This is either the folder that the user chose or the default shown on the GUI
            'filename = "C:\Data\" & logfiletextbox.Text & ".csv"
            FileOpen(1, filename, OpenMode.Append, OpenAccess.Write)
            WriteLine(1, DateTime.Now(), temperature, tbsetPoint.Text)
            FileClose(1)
        Catch ex As Exception
            'file already open
            logfiletextbox.Text += "2"
        End Try

    End Sub

    Public Sub hover(ByVal hovert As Double)
        'This sub takes in an amount of time, in minutes, and loops until that amount of time has elapsed

        Dim remaining As TimeSpan
        Dim delay_time As Date
        Dim t_remain As Integer

        delay_time = TimeValue(Now)
        Do
            remaining = TimeValue(delay_time.AddMinutes(hovert)) - TimeValue(Now)
            t_remain = remaining.Seconds + remaining.Minutes * 60 + remaining.Hours * 3600
            updatetextbox.Text = "Holding; " + Convert.ToString(t_remain) + " s remain"

            If TimeValue(Now) >= TimeValue(delay_time.AddMinutes(hovert)) Then
                Exit Do
            End If

            currentTemp.Text = read_eurotemp()
            Threading.Thread.Sleep(5)
            Application.DoEvents()
        Loop

    End Sub

    Public Sub rampup(ByVal t0 As Double, ByVal t1 As Double, ByVal rate As Double)
        'This function handles only ramping up in temperature
        'Determine the step size, then how many iterations at that step to get from t0 -> t1
        'Then step up the set temperature and wait one minute.

        Dim diff, increments, temp_step As Double

        diff = Math.Abs(t1 - t0)
        increments = Math.Round(diff / rate)

        For k As Integer = 1 To increments

            temp_step = t0 + k * rate
            set_eurotemp(temp_step)     'set_eurotemp only accepts integers...this means I have to look at how to do this.
            '                           ' Check later to see if I can get this working by changing some of the set_eurotemp
            '                           ' code working. 

            hover(1)
        Next
    End Sub

    Public Sub rampdown(ByVal t0 As Double, ByVal t1 As Double, ByVal rate As Double)
        'This function handles only ramping down in temperature
        'Determine the step size, then how many iterations at that step to get from t0 -> t1
        'Then step up the set temperature and wait one minute.

        Dim diff, increments, temp_step As Double

        diff = Math.Abs(t1 - t0)
        increments = Math.Round(diff / rate)

        For k As Integer = 1 To increments

            temp_step = t0 - k * rate
            set_eurotemp(temp_step)

            hover(1)
        Next

    End Sub

    Function precheckTemps() As Integer
        'Goes through each text box and explicitly checks for temperature mismatches
        'So, if ramp up is chosen, we need T0 < T1
        'Likewise, for ramp down, we need  T0 > T1
        Dim t0, t1, rate, hover_t As Double
        Dim listbox_list() As ListBox
        Dim textboxarray() As TextBox

        listbox_list = {ListBox1, ListBox2, ListBox3, ListBox4, ListBox5, ListBox6, ListBox7, ListBox8, ListBox9}
        textboxarray = {TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12, TextBox13, TextBox14,
            TextBox15, TextBox16, TextBox17, TextBox18, TextBox19, TextBox20, TextBox21, TextBox22, TextBox23, TextBox24, TextBox25, TextBox26, TextBox27,
            TextBox28, TextBox29, TextBox30, TextBox31, TextBox32, TextBox33, TextBox34, TextBox35, TextBox36, TextBox37}

        For ii As Integer = 0 To listbox_list.Length - 1
            'Loop over all the listboxes to see what the selection is.
            'If no selection is made, or it is null, then do nothing and continue the loop.

            If listbox_list(ii).SelectedIndex = -1 Then 'No selection
                Continue For
            End If
            If listbox_list(ii).SelectedIndex = 0 Then 'Blank selection
                Continue For
            End If

            If listbox_list(ii).SelectedIndex = 1 Then  'Ramp up

                t0 = 2 + (ii * 4) - 1
                t0 = Convert.ToDouble(textboxarray(t0).Text)
                t1 = 3 + (ii * 4) - 1
                t1 = Convert.ToDouble(textboxarray(t1).Text)
                rate = 4 + (ii * 4) - 1
                rate = Convert.ToDouble(textboxarray(rate).Text)

                If t0 < t1 = False Then
                    Return 0 'On fail, return 0
                End If
                If rate <= 0 Then
                    Return 0 'On fail, return 0
                End If
            ElseIf listbox_list(ii).SelectedIndex = 2 Then  'ramp down

                t0 = 2 + (ii * 4) - 1
                t0 = Convert.ToDouble(textboxarray(t0).Text)
                t1 = 3 + (ii * 4) - 1
                t1 = Convert.ToDouble(textboxarray(t1).Text)
                rate = 4 + (ii * 4) - 1
                rate = Convert.ToDouble(textboxarray(rate).Text)

                If t0 > t1 = False Then
                    Return 0
                End If
                If rate <= 0 Then
                    Return 0 'On fail, return 0
                End If
            ElseIf listbox_list(ii).SelectedIndex = 3 Then  'hover

                hover_t = 5 + (ii * 4) - 1
                hover_t = Convert.ToDouble(textboxarray(hover_t).Text)

                If hover_t <= 0 Then
                    Return 0
                End If
            End If
        Next

        Return 1 'On success return 1

    End Function

    Private Sub singleSetTempButton_Click(sender As Object, e As EventArgs) Handles singleSetTempButton.Click

        Dim gototemp As Double

        updatetextbox.Text = "Moving to T = " & singleSetTemp.Text
        gototemp = singleSetTemp.Text
        set_eurotemp(gototemp)
        Application.DoEvents()

    End Sub

    Private Sub setPathButton_Click(sender As Object, e As EventArgs) Handles setPathButton.Click

        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()

        If (result = DialogResult.OK) Then
            updatetextbox.Text = "Folder selected is: " & ControlChars.NewLine & FolderBrowserDialog1.SelectedPath
            Application.DoEvents()
        End If

    End Sub
End Class
