Imports System.Security.Cryptography.X509Certificates

Public Class Form1
    '
    '  Realizar validaciones de datos introducidos
    '
    Private cbColumnAccount, cbColumnGrupoMedio As New DataGridViewComboBoxColumn
    Private conexion As AccesoDB = New AccesoDB()

    Public Sub New()
        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        anyoActual.HeaderText = System.DateTime.Now.Year
        anyoAnterior.HeaderText = System.DateTime.Now.Year - 1
        cbColumnAccount = conexion.GetDataComboBoxColumn("Account", "Account")
        cbColumnGrupoMedio = conexion.GetDataComboBoxColumn("CfgGrupoMedio", "Grupo Medio")
        dgvClientDatabase.ColumnHeadersHeight = 90
        dgvClientDatabase.Columns.Insert(0, cbColumnAccount)
        dgvClientDatabase.Columns.Insert(1, cbColumnGrupoMedio)
    End Sub

    Private Sub btnInsetar_Click(sender As Object, e As EventArgs) Handles btnInsetar.Click

        Dim prevision As New Prevision()
        Dim fila As Integer = dgvClientDatabase.SelectedCells(0).RowIndex
        Dim columna As Integer = dgvClientDatabase.SelectedCells(0).ColumnIndex
        Dim cell As String = dgvClientDatabase.Rows(fila).Cells(columna).Value
        Dim valorString As String = dgvClientDatabase.Rows(fila).Cells(15).Value
        Dim valorDouble As Double
        '
        ' Evitar pasar directamente el contenido de un control a
        ' cualquier procedimiento. Antes hay que validar totalmente el conteniod
        ' de los controles y almacenarlos en variables locales.
        '
        prevision._grupoMedioId = 1
        prevision._clientDatabaseId = 4
        prevision._mesId = columna
        If Double.TryParse(valorString, valorDouble) Then
            prevision._importe = valorDouble
        End If
        conexion.InsertarPrevision(prevision)
        'MsgBox("Fila en la que está el foco: " & fila & " Columna: " & columna & ". Valor: " & cell)
    End Sub


    Private Sub dgvClientDatabase_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvClientDatabase.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.V Then
            'PegarDelPortapapeles()
        End If
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        'If dgvClientDatabase.Rows.Count > 0 Then
        '    For fila As Integer = 0 To dgvClientDatabase.Rows.Count
        '        Dim cont As String = dgvClientDatabase.Rows(fila).Cells(1).Value
        '        MsgBox("Fila: " & cont)
        '    Next
        'End If
        'Dim filasSeleccionadas As DataGridViewSelectedRowCollection = dgvClientDatabase.SelectedRows()
        'MessageBox.Show("Sub Eliminar registro")
    End Sub


    Private Sub dgvClientDatabase_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvClientDatabase.CellEnter
        Dim sumaAnual As Double = 0
        Dim nFila As Integer

        ' Comprueba que al menos una celda ha obtenido el foco
        If dgvClientDatabase.SelectedCells.Count > 0 Then

            nFila = dgvClientDatabase.SelectedCells(0).RowIndex
            If Not nFila Mod 2 = 0 Then
                dgvClientDatabase.Rows(nFila).DefaultCellStyle.BackColor = Color.Coral
            End If

            dgvClientDatabase.Rows(nFila).Cells(15).Value = "0.0"

            ' Recorre todas las celdas de la fila actual
            For Each celda As DataGridViewCell In dgvClientDatabase.Rows(nFila).Cells
                ' Acota entre las columnas 2 y 13 (Las que corresponden a los meses)
                If e.ColumnIndex > 1 And e.ColumnIndex < 14 Then
                    For i As Integer = 2 To 13
                        ' Inserta el valor 0.0 a todas las celdas excepto la seleccionada
                        ' y las que el valor sea mayor de 0
                        '
                        ' TO DO: El If hay que revisarlo
                        Dim valorCelda As String = dgvClientDatabase.Rows(nFila).Cells(i).Value
                        Dim valorDouble As Double = Double.TryParse(valorCelda, valorDouble)

                        If i <> e.ColumnIndex And valorDouble = 0 Then
                            dgvClientDatabase.Rows(nFila).Cells(i).Value = "0.0"
                        End If
                    Next
                End If
            Next

            ' Realiza la suma de las cantidades de todos los meses
            For i As Integer = 2 To 13
                Dim valorString As String = dgvClientDatabase.Rows(nFila).Cells(i).Value
                Dim valorDouble As Double

                If Double.TryParse(valorString, valorDouble) Then
                    sumaAnual += valorDouble
                End If
            Next
            '  MsgBox("Total anual: " & sumaAnual)
            dgvClientDatabase.Rows(nFila).Cells(15).Value = sumaAnual
        End If
    End Sub

    ' CellValidating sólo lanza el mensaje de error la primera vez que se introduce
    ' mal el dato. El resto de las veces simplemente inicia la celda a 0,0
    ' 
    'Private Sub dgvClientDatabase_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvClientDatabase.CellValidating
    '    If e.ColumnIndex > 1 Then
    '        If Not IsNumeric(dgvClientDatabase.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
    '            MsgBox("El valor introducido no es correcto.")
    '        End If
    '    End If
    'End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Rellenar el DataGridView con los datos que ya se hayan guardado en
        ' la DB
    End Sub

    ' CellValidated lanza mensaje de error siempre que la casilla 
    ' obtenga un dato erróneo
    Private Sub dgvClientDatabase_CellValidated(sender As Object, e As DataGridViewCellEventArgs) Handles dgvClientDatabase.CellValidated
        If e.ColumnIndex > 1 Then
            If Not IsNumeric(dgvClientDatabase.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                MsgBox("El valor introducido no es correcto.")
            End If
        End If
    End Sub
End Class
