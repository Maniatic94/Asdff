Modificaciones al Formulario
Agregar el TextBox:

Abre el formulario en el editor de diseño.
Agrega un nuevo control TextBox y nómbralo txtTecla.
Coloca un Label para identificar el TextBox, como "Tecla Asociada:".
Modificar el Código para Manejar la Tecla:

Declaraciones Globales
Añade un campo Key en la estructura de macro para almacenar la tecla configurada

Public Type tMacros
    mTipe As Byte
    grh As Long
    nombre As String
    Slot As Byte
    OBJIndex As Integer
    SpellSlot As Byte
    Key As String ' Nueva propiedad para almacenar la tecla
End Type

Al Guardar la Configuración
En el evento cmdAccept_Click, guarda la tecla introducida en el TextBox:

Private Sub cmdAccept_Click()
    On Error Resume Next

    Dim i As Integer
    For i = optAccion.LBound To optAccion.UBound
        If optAccion(i).Value = True Then
            MacroList(MacroIndex).mTipe = i + 1
            Exit For
        End If
    Next i

    ' Guardar la tecla configurada
    MacroList(MacroIndex).Key = txtTecla.Text

    ' Lógica existente para cada tipo de macro
    Select Case MacroList(MacroIndex).mTipe
        Case eMacros.aComando
            If LenB(Text1.Text) = 0 Then
                MacroList(MacroIndex).mTipe = 0
                Exit Sub
            End If
            MacroList(MacroIndex).mTipe = eMacros.aComando
            MacroList(MacroIndex).grh = 17506
            MacroList(MacroIndex).nombre = UCase$(Text1.Text)
        Case eMacros.aLanzar
            If hlst.List(hlst.ListIndex) = "(Vacio)" Then
                MacroList(MacroIndex).mTipe = 0
                Exit Sub
                Unload Me
            End If
            MacroList(MacroIndex).mTipe = eMacros.aLanzar
            MacroList(MacroIndex).grh = 609
            MacroList(MacroIndex).nombre = hlst.List(hlst.ListIndex)
            MacroList(MacroIndex).SpellSlot = hlst.List(hlst.ListIndex) + 1
        Case eMacros.aUsar
            If frmMain.Inventario.SelectedItem = 0 Then
                MacroList(MacroIndex).mTipe = 0
                Unload Me
                Exit Sub
            End If
            MacroList(MacroIndex).mTipe = eMacros.aUsar
            MacroList(MacroIndex).grh = frmMain.Inventario.GrhIndex(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).nombre = frmMain.Inventario.ItemName(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).OBJIndex = frmMain.Inventario.OBJIndex(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).Slot = frmMain.Inventario.SelectedItem
        Case eMacros.aEquipar
            If frmMain.Inventario.SelectedItem = 0 Then
                MacroList(MacroIndex).mTipe = 0
                Unload Me
                Exit Sub
            End If
            MacroList(MacroIndex).mTipe = eMacros.aEquipar
            MacroList(MacroIndex).grh = frmMain.Inventario.GrhIndex(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).nombre = frmMain.Inventario.ItemName(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).OBJIndex = frmMain.Inventario.OBJIndex(frmMain.Inventario.SelectedItem)
            MacroList(MacroIndex).Slot = frmMain.Inventario.SelectedItem
    End Select

    ' Guardar las macros en el archivo
    Call SaveMacros(frmMain.NombrePJ)
    ' Actualizar el gráfico del macro
    Call Grh_Render_To_Hdc(frmMain.picMacro(MacroIndex), MacroList(MacroIndex).grh, 0, 0)

    Unload Me
End Sub

Mostrar la Tecla Configurada
En el evento Form_Load, muestra la tecla configurada en el TextBox:

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

    If MacroList(MacroIndex).mTipe <> 0 Then
        Select Case MacroList(MacroIndex).mTipe
            Case eMacros.aComando
                optAccion(1).Value = True
                Text1.Text = MacroList(MacroIndex).nombre
                Text1.Enabled = True
        End Select
    End If

    ' Mostrar la tecla configurada
    txtTecla.Text = MacroList(MacroIndex).Key
End Sub

Manejo de Teclas Asociadas
Para ejecutar los macros según las teclas configuradas:

Agrega un manejador de eventos en el formulario principal para capturar las teclas presionadas.
Compara la tecla presionada con la configuración de cada macro.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    For i = 1 To 5
        If MacroList(i).Key = Chr$(KeyCode) Then
            Call UsarMacro(i)
            Exit For
        End If
    Next i
End Sub

Recuerda establecer la propiedad KeyPreview del formulario principal a True.

Guardar y Cargar Teclas
Asegúrate de modificar las funciones SaveMacros y LoadMacros para incluir la propiedad Key:

Public Sub SaveMacros(ByVal nombre As String)
    Dim MacroPatch As String
    Dim i As Integer
    MacroPatch = App.Path & "\..\Recursos\Macros\" & nombre & ".Mac"

    For i = 1 To 5
        With MacroList(i)
            Call WriteVar(MacroPatch, "Macro" & i, "Nombre", .nombre)
            Call WriteVar(MacroPatch, "Macro" & i, "Grh", .grh)
            Call WriteVar(MacroPatch, "Macro" & i, "Tipo", .mTipe)
            Call WriteVar(MacroPatch, "Macro" & i, "Slot", .Slot)
            Call WriteVar(MacroPatch, "Macro" & i, "SlotSpell", .SpellSlot)
            Call WriteVar(MacroPatch, "Macro" & i, "ObjIndex", .OBJIndex)
            Call WriteVar(MacroPatch, "Macro" & i, "Key", .Key) ' Guardar la tecla
        End With
    Next i
End Sub

De esta forma, cada macro tiene una tecla asociada que puede ser configurada y usada de manera independiente.
