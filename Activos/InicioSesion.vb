Imports System.Runtime.InteropServices

Public Class InicioSesion

    Private Sub CustomizeComponents()
        'txtUser
        txtUser.AutoSize = False
        txtUser.Size = New Size(350, 30)
        'txtPass
        txtPass.AutoSize = False
        txtPass.Size = New Size(350, 30)
        txtPass.UseSystemPasswordChar = True
    End Sub

    Private Sub btnLogin_Paint(sender As Object, e As PaintEventArgs) Handles btnLogin.Paint
        Dim buttonPath As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath()
        Dim myRectangle As Rectangle = btnLogin.ClientRectangle
        myRectangle.Inflate(0, 30)
        buttonPath.AddEllipse(myRectangle)
        btnLogin.Region = New Region(buttonPath)
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub
    Private Sub btnMinimize_Click(sender As Object, e As EventArgs) Handles btnMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

#Region "Drag Form - Arrastrar/ mover Formulario"

    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub
    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(hWnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer)
    End Sub

    Private Sub titleBar_MouseDown(sender As Object, e As MouseEventArgs) Handles titleBar.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomizeComponents()
    End Sub

    ''' <summary>
    ''' Validar el inicio de sesión contra la BD
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If Not txtUser.Text.Equals("") And Not txtPass.Text.Equals("") Then
            Dim Conexion As SQLite.SQLiteConnection
            Dim Adaptador As SQLite.SQLiteDataAdapter

            Conexion = New SQLite.SQLiteConnection
            Conexion.ConnectionString = "Data Source='C:\Proyectos\Activos Escritorio\Activos_Escritorio\Activos\bin\Debug\activos.db';Version=3;"

            Conexion.Open()

            Dim ds As New DataSet
            Dim sql As String
            sql = String.Format("SELECT * FROM USUARIO WHERE usuario = '{0}' AND contrasena = '{1}'", txtUser.Text, txtPass.Text)
            Adaptador = New SQLite.SQLiteDataAdapter(sql, Conexion)
            Adaptador.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                idUsuario = Integer.Parse(ds.Tables(0).Rows(0)(0).ToString())
                cedula = ds.Tables(0).Rows(0)(2).ToString()
                nombre = ds.Tables(0).Rows(0)(3).ToString() + " " + ds.Tables(0).Rows(0)(4).ToString()
                Conexion.Close()
                Me.Hide()
                FormPrincipal.Show()
            Else
                MsgBox("El usuario no existe.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Inicio de Sesión")
            End If
        Else
            MsgBox("Por favor ingrese los datos.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Inicio de Sesión")
        End If
    End Sub
End Class
