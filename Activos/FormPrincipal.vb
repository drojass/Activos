Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Status

Public Class FormPrincipal
#Region "FUNCIONALIDADES DEL FORMULARIO"
    'RESIZE DEL FORMULARIO- CAMBIAR TAMAÑO
    Dim cGrip As Integer = 10

    Protected Overrides Sub WndProc(ByRef m As Message)
        If (m.Msg = 132) Then
            Dim pos As Point = New Point((m.LParam.ToInt32 And 65535), (m.LParam.ToInt32 + 16))
            pos = Me.PointToClient(pos)
            If ((pos.X _
                        >= (Me.ClientSize.Width - cGrip)) _
                        AndAlso (pos.Y _
                        >= (Me.ClientSize.Height - cGrip))) Then
                m.Result = CType(17, IntPtr)
                Return
            End If
        End If
        MyBase.WndProc(m)
    End Sub
    '----------------DIBUJAR RECTANGULO / EXCLUIR ESQUINA PANEL 
    Dim sizeGripRectangle As Rectangle
    Dim tolerance As Integer = 15

    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        MyBase.OnSizeChanged(e)
        Dim region = New Region(New Rectangle(0, 0, Me.ClientRectangle.Width, Me.ClientRectangle.Height))
        sizeGripRectangle = New Rectangle((Me.ClientRectangle.Width - tolerance), (Me.ClientRectangle.Height - tolerance), tolerance, tolerance)
        region.Exclude(sizeGripRectangle)
        Me.PanelContenedor.Region = region
        Me.Invalidate()
    End Sub

    '----------------COLOR Y GRIP DE RECTANGULO INFERIOR
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim blueBrush As SolidBrush = New SolidBrush(Color.FromArgb(244, 244, 244))
        e.Graphics.FillRectangle(blueBrush, sizeGripRectangle)
        MyBase.OnPaint(e)
        ControlPaint.DrawSizeGrip(e.Graphics, Color.Transparent, sizeGripRectangle)
    End Sub
    'ARRASTRAR EL FORMULARIO
    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub

    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    End Sub

    Private Sub PanelBarraTitulo_MouseMove(sender As Object, e As MouseEventArgs) Handles PanelBarraTitulo.MouseMove
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Application.Exit()
    End Sub

    Dim lx, ly As Integer
    Dim sw, sh As Integer

    Private Sub btnMinimizar_Click(sender As Object, e As EventArgs) Handles btnMinimizar.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub btnMaximizar_Click(sender As Object, e As EventArgs) Handles btnMaximizar.Click

        lx = Me.Location.X
        ly = Me.Location.Y
        sw = Me.Size.Width
        sh = Me.Size.Height

        btnMaximizar.Visible = False
        btnRestaurar.Visible = True

        Me.Size = Screen.PrimaryScreen.WorkingArea.Size
        Me.Location = Screen.PrimaryScreen.WorkingArea.Location

    End Sub



    Private Sub btnRestaurar_Click(sender As Object, e As EventArgs) Handles btnRestaurar.Click
        Me.Size = New Size(sw, sh)
        Me.Location = New Point(lx, ly)

        btnMaximizar.Visible = True
        btnRestaurar.Visible = False
    End Sub
#End Region

    'METODO DE ABRIR FORMULARIO
    Private Sub AbrirFormEnPanel(Of Miform As {Form, New})()
        Dim Formulario As Form
        Formulario = PanelFormularios.Controls.OfType(Of Miform)().FirstOrDefault() 'Busca el formulario en la coleccion
        'Si form no fue econtrado/ no existe
        If Formulario Is Nothing Then
            Formulario = New Miform With {
                .TopLevel = False,
                .FormBorderStyle = FormBorderStyle.None,
                .Dock = DockStyle.Fill
            }

            PanelFormularios.Controls.Add(Formulario)
            PanelFormularios.Tag = Formulario
            AddHandler Formulario.FormClosed, AddressOf Me.CerrarFormulario
            Formulario.Show()
            Formulario.BringToFront()
        Else
            Formulario.BringToFront()
        End If

    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs)
    '    AbrirFormEnPanel(Of Form1)()
    '    Button1.BackColor = Color.FromArgb(12, 61, 92)
    'End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs)
    '    AbrirFormEnPanel(Of Form2)()
    '    Button2.BackColor = Color.FromArgb(12, 61, 92)
    'End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        logo2.Visible = False
        logo.Visible = True
        Timer2.Enabled = True
        Timer1.Enabled = False
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        logo.Visible = False
        logo2.Visible = True
        Timer1.Enabled = True
        Timer2.Enabled = False
    End Sub

    Private Sub FormPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer2.Enabled = True
        lblBienvenido.Text = "Bienvenido " + nombre
        Dim Conexion As SQLite.SQLiteConnection
        Dim Adaptador As SQLite.SQLiteDataAdapter

        Conexion = New SQLite.SQLiteConnection
        Conexion.ConnectionString = "Data Source='C:\Proyectos\Activos Escritorio\Activos_Escritorio\Activos\bin\Debug\activos.db';Version=3;"

        Conexion.Open()

        Dim ds As New DataSet
        Dim sql As String
        sql = String.Format("SELECT m.*
                            FROM MENU m 
                            INNER JOIN ROL_MENU rm 
                            ON m.menu_id = rm.menu_id 
                            INNER JOIN ROL r 
                            ON rm.rol_id = r.rol_id
                            INNER JOIN USUARIO_ROL ur 
                            ON r.rol_id = ur.persona 
                            WHERE ur.persona = {0};", idUsuario)
        Adaptador = New SQLite.SQLiteDataAdapter(sql, Conexion)
        Adaptador.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            For I As Integer = 0 To ds.Tables(0).Rows.Count - 1
                Dim b As New Button()
                b.Name = "btn" + ds.Tables(0).Rows(I)(2).ToString()
                b.Text = ds.Tables(0).Rows(I)(2).ToString()
                b.Height = 32
                b.Width = 188
                b.BackColor = Color.FromArgb(48, 63, 105)
                b.ForeColor = Color.WhiteSmoke
                'Opcional, conectar el evento click:
                AddHandler b.Click, AddressOf Btn_Click
                PanelMenu.Controls.Add(b)
            Next I
        End If
    End Sub

    Private Sub Btn_Click(sender As Object, e As EventArgs)
        MsgBox("Acción del bóton.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Acción del bóton.")
    End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs)
    '    AbrirFormEnPanel(Of Form3)()
    '    Button3.BackColor = Color.FromArgb(12, 61, 92)
    'End Sub
    'METODO/EVENTO AL CERRAR FORMS
    Private Sub CerrarFormulario(ByVal sender As Object, ByVal e As FormClosedEventArgs)
        'CONDICION SI FORMS ESTA ABIERTO
        'If (Application.OpenForms("Form1") Is Nothing) Then
        '    Button1.BackColor = Color.FromArgb(4, 41, 68)
        'End If
        'If (Application.OpenForms("Form2") Is Nothing) Then
        '    Button2.BackColor = Color.FromArgb(4, 41, 68)
        'End If
        'If (Application.OpenForms("Form3") Is Nothing) Then
        '    Button3.BackColor = Color.FromArgb(4, 41, 68)
        'End If
    End Sub
End Class
