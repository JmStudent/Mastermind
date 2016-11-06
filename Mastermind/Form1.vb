Public Class Form1
    Dim menuColors As New ContextMenuStrip
    Dim btnSelected As New Button
    Dim tabBtn(9, 4) As Button
    Dim tabBtnWin As New ArrayList
    Dim tabLblBW(9, 1) As Label
    Dim posR As Integer ' = 0
    Dim activeRow As New ArrayList
    Dim btnCheck As New Button
    Dim color As String() = {"Red", "Orange", "Yellow", "YellowGreen", "Green", "DeepSkyBlue", "Purple", "DarkGray"}
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ContextMenuStrip1.Items.Add(New ToolStripControlHost(New Label With {.Text = "Title"}))
        'ContextMenuStrip1.Items.Add(Label1.Text, Nothing, AddressOf optionColorsMouseClick)
        ' creación del menú de opciones
        For i = 0 To 7
            Dim optionColor As New ToolStripMenuItem(color(i))
            Dim path As String = Application.StartupPath & "\img\" & i & ".png"
            optionColor.Image = Image.FromFile(path)
            optionColor.BackColor = System.Drawing.Color.FromName(color(i))
            optionColor.Tag = i
            AddHandler optionColor.MouseUp, AddressOf optionColors_MouseClick
            menuColors.Items.Add(optionColor)
        Next
        Dim optionWithoutColor As New ToolStripMenuItem("Sin color")
        optionWithoutColor.Tag = 9
        AddHandler optionWithoutColor.MouseUp, AddressOf optionColors_MouseClick
        menuColors.Items.Add(optionWithoutColor)

        ' creación de la botonera
        For i = 0 To 9
            For j = 0 To 4
                Dim btn As New Button
                btn.Size = New Drawing.Size(45, 45)
                btn.Location = New Drawing.Point(25 + (45 * j), 150 + (45 * i))
                btn.Name = "btn" & i & j
                btn.Tag = 9 ' el 9 significa que no tiene color asignado
                btn.Enabled = False
                tabBtn(i, j) = btn
                Me.Controls.Add(btn)
                AddHandler btn.MouseUp, AddressOf btn_MouseClick
            Next
        Next

        'creación del botón comprobar
        btnCheck.Size = New Drawing.Size(120, 45)
        btnCheck.Location = New Drawing.Point(115, 12)
        btnCheck.Font = New Font("Arial", 12)
        btnCheck.Text = "Check"
        btnCheck.Name = "btnCheck"
        btnCheck.Enabled = False
        Me.Controls.Add(btnCheck)
        AddHandler btnCheck.MouseUp, AddressOf btnCheck_MouseClick

        ' creación de la fila ganadora
        For j = 0 To 4
            Dim btnWin As New Button
            btnWin.Size = New Drawing.Size(45, 45)
            btnWin.Location = New Drawing.Point(25 + (45 * j), 85)
            btnWin.Name = "btnWin" & j
            btnWin.Tag = randNum()
            btnWin.Enabled = False
            tabBtnWin.Add(btnWin)
            Me.Controls.Add(btnWin)
        Next

        ' creación de las etiquetas que tendrán lo de blanco y negro
        ' Aquí las dos de ejemplo
        For j = 0 To 1
            Dim lblnb As New Label
            lblnb.Size = New Drawing.Size(45, 45)
            lblnb.Location = New Drawing.Point(255 + (45 * j), 85)
            lblnb.BorderStyle = BorderStyle.FixedSingle
            lblnb.TextAlign = ContentAlignment.MiddleCenter
            lblnb.Font = New Font("Arial", 15, FontStyle.Bold)
            If j Mod 2 = 0 Then
                lblnb.BackColor = System.Drawing.Color.FromName("black")
                lblnb.ForeColor = System.Drawing.Color.FromName("white")
                lblnb.Text = "N"
            Else
                lblnb.BackColor = System.Drawing.Color.FromName("white")
                lblnb.ForeColor = System.Drawing.Color.FromName("black")
                lblnb.Text = "B"
            End If
            Me.Controls.Add(lblnb)
        Next

        ' y aquí todo el resto de etiquetas
        For i = 0 To 9
            For j = 0 To 1
                Dim lblBW As New Label
                lblBW.Size = New Drawing.Size(45, 45)
                lblBW.Location = New Drawing.Point(255 + (45 * j), 150 + (45 * i))
                lblBW.Name = "lblBW" & i & j
                lblBW.BorderStyle = BorderStyle.FixedSingle
                lblBW.TextAlign = ContentAlignment.MiddleCenter
                lblBW.Font = New Font("Arial", 15, FontStyle.Bold)
                If j Mod 2 = 0 Then
                    lblBW.BackColor = System.Drawing.Color.FromName("black")
                    lblBW.ForeColor = System.Drawing.Color.FromName("white")
                Else
                    lblBW.BackColor = System.Drawing.Color.FromName("white")
                    lblBW.ForeColor = System.Drawing.Color.FromName("black")
                End If
                tabLblBW(i, j) = lblBW
                lblBW.Visible = False
                Me.Controls.Add(lblBW)
            Next
        Next
        ' coloring()
        unlock() ' desbloquea la primera fila
    End Sub

    ' MANEJADORES DE EVENTOS
    ' MANEJADOR CUANDO LE DAMOS CLICK A UN BOTÓN
    Private Sub btn_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) ' Handles Me.MouseUp
        ' cada vez que le doy a un botón hace lo siguiente:
        If sender.GetType.ToString = "System.Windows.Forms.Button" Then ' si verdaderamente lo que he pulsado es un botón
            menuColors.Show(sender, 0, sender.Height) ' asocia el menú al botón pulsado
            changeSelected(sender) ' definimos el atributo btnSelected como el objeto marcado
        End If
    End Sub

    'MANEJADOR DEL MENÚ DE OPCIONES CON LOS COLORES
    Private Sub optionColors_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) ' Handles Me.MouseUp
        If sender.GetType.ToString = "System.Windows.Forms.ToolStripMenuItem" Then ' si el evento viene del menú
            Dim permission As Boolean = True
            For Each btnInRow In activeRow ' este for each evita que se puede poner dos veces el mismo color
                If btnInRow.Tag = sender.Tag Then
                    If btnInRow.Tag <> 9 Then
                        permission = False
                        Exit For
                    End If
                End If
            Next
            If permission Then
                btnSelected.BackColor = sender.BackColor ' ponerle el mismo color de fondo que el del menú
                btnSelected.Tag = sender.Tag ' y el mismo tag
            End If
            If checkActiveRowFull() Then
                btnCheck.Enabled = True
            Else
                btnCheck.Enabled = False
            End If
        End If
    End Sub
    ' MANEJADOR BOTON CHECK
    Private Sub btnCheck_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs)
        ' Falta
        ' comprobar la combinación activa con la ganadora
        ' para ello recorremos la activa y vamos comprobando cada uno, cada check positivo suma 1 y si dan 5 checks devuelve true
        ' si ganas se muestra un mensaje y se muestra la combinación ganadora
        ' si no pues se desbloquea la siguiente y ya está.
        ' todo esto deberá ir dentro de checkcompleterow para asegurarnos que estas acciones se hacen con toda la fila completa
        Dim black As Integer = checkBlacks()
        showBlackWhite(black)
        If black = 5 Then
            coloring()
            msgEnd("Has ganado!!!")
        ElseIf posR = 10 Then
            coloring()
            msgEnd("Has perdido")
        End If
        ' además para controlar que pierde, la condición será si pulsas check y la posR == 9 (se debe poner después de la de ganar)
        ' Se muestra un mensaje y se dice que ha perdido
        If checkCompleteRow() Then
            lock()
            unlock()
        End If
        btnCheck.Enabled = False
    End Sub
    Private Sub changeSelected(btn As Button)
        btnSelected = btn
    End Sub
    Private Sub lock()
        If posR < 10 Then
            For i = 0 To posR
                For j = 0 To 4
                    tabBtn(i, j).Enabled = False
                Next
            Next
        End If
    End Sub
    Private Sub unlock()
        If posR < 10 Then
            For j = 0 To 4
                tabBtn(posR, j).Enabled = True
                activeRow.Add(tabBtn(posR, j))
            Next
        End If
        posR += 1
    End Sub
    Private Function checkCompleteRow()
        Dim suma As Integer
        For Each x In activeRow
            If x.Tag <> 9 Then
                suma += 1
            End If
        Next
        If suma = 5 Then
            ' activeRow.RemoveRange(0, 5)  -> hace lo mismo que clear()
            activeRow.Clear()
        End If
        Return suma = 5
    End Function
    Private Function randNum()
        Randomize()
        ' método que devuelve un número aleatorio. randNum()
        ' el método crea un número aleatorio, si arraylist está vacio, lo mete del tiron, sino, comprueba que no esté y si está crea uno nuevo hasta que no esté.
        ' btnWin.BackColor = System.Drawing.Color.FromName("Red")
        Dim rand As Integer
        Do
            rand = Int(7 * Rnd())
            If tabBtnWin.Count = 0 Then
                Return rand
            End If
        Loop While isRepeat(rand)
        Return rand
    End Function
    Private Function isRepeat(num As Integer)
        For Each btn In tabBtnWin
            If btn.Tag = num Then
                Return True ' si está repetido
            End If
        Next
        Return False ' si no lo está
    End Function
    Private Sub coloring()
        For Each btn In tabBtnWin
            btn.BackColor = System.Drawing.Color.FromName(color(btn.Tag))
        Next
    End Sub
    Private Sub showBlackWhite(black As Integer)
        ''Dim black As String
        'black = checkBlacks()
        Dim pos As Integer
        ' MsgBox(posR)
        pos = posR - 1
        'If posR = 9 Then
        '    pos = posR
        'Else
        '    pos = posR - 1
        'End If
        tabLblBW(pos, 0).Text = black ' matchesBlack.Count
        tabLblBW(pos, 0).Visible = True
        tabLblBW(pos, 1).Text = checkWhites() - black
        tabLblBW(pos, 1).Visible = True
        ' y aquí asignar el white y el black al texto de las etiquetas.
        ' llamar a este método tras darle a comprobar y sólo si no se ha ganado.
    End Sub
    Private Function checkBlacks()
        Dim black As Integer = 0
        For i = 0 To activeRow.Count - 1
            If tabBtnWin.Item(i).Tag = activeRow.Item(i).Tag Then
                'matchesBlack.Add(activeRow.Item(i).Tag)
                black += 1
            End If
        Next
        Return black
    End Function
    Private Function checkWhites()
        Dim white As Integer = 0
        For Each btnActive In activeRow
            For Each btnW In tabBtnWin
                If btnActive.Tag = btnW.Tag Then
                    white += 1
                End If
            Next
        Next
        Return white
    End Function
    Private Sub msgEnd(msg As String)
        Dim res As Integer
        res = MsgBox(msg + " ¿Desea reiniciar?", vbYesNo)
        If res = 6 Then
            Application.Restart()
        Else
            Application.Exit()
        End If
    End Sub
    Private Function checkActiveRowFull()
        For Each btn In activeRow
            If btn.Tag = 9 Then
                Return False
            End If
        Next
        Return True
    End Function
End Class
