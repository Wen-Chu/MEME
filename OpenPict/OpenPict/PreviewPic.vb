Imports System.IO
Imports System.Data.SqlClient

Public Class PreviewPic

    Dim connection As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\VB\OpenPict\OpenPict\Picture.mdf;Integrated Security=True")
    Dim ssub As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\VB\OpenPict\OpenPict\Picture.mdf;Integrated Security=True")
    Dim id_count As New SqlCommand("select time from Picture", connection)
    Dim table As New DataTable()
    Dim adapter As New SqlDataAdapter(id_count)
    Dim count As Int32 = adapter.Fill(table)

    Dim command As New SqlCommand("Select * from Picture where id = @id", connection)

    Public Sub PreviewPic_Show(ByVal num As Integer)
        Select Case num
            Case 1
                PictureBox1.Image = MainForm.PictureBox1.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count
                Label1.Text = getTime().ToString
            Case 2
                PictureBox1.Image = MainForm.PictureBox2.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 1
                Label1.Text = getTime().ToString
            Case 3
                PictureBox1.Image = MainForm.PictureBox3.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 2
                Label1.Text = getTime().ToString
            Case 4
                PictureBox1.Image = MainForm.PictureBox4.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 3
                Label1.Text = getTime().ToString
            Case 5
                PictureBox1.Image = MainForm.PictureBox5.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 4
                Label1.Text = getTime().ToString
            Case 6
                PictureBox1.Image = MainForm.PictureBox6.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 5
                Label1.Text = getTime().ToString
            Case 7
                PictureBox1.Image = MainForm.PictureBox7.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 6
                Label1.Text = getTime().ToString
            Case 8
                PictureBox1.Image = MainForm.PictureBox8.Image
                command.Parameters.Add("@id", SqlDbType.Int).Value = count - 7
                Label1.Text = getTime().ToString
        End Select
    End Sub

    Function getTime()
        table = New DataTable()
        adapter = New SqlDataAdapter(Command)

        adapter.Fill(table)
        Dim time As DateTime = table.Rows(0)(2)
        Return time
    End Function

    Private Sub Preview_Picture_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Font = New Font("Lucida Fax", Label1.Font.Size, Label1.Font.Style)
        Label1.Location = New Point(0, PictureBox1.Height)
        Label1.Width = PictureBox1.Width
        Me.Width = PictureBox1.Width + 15
        Me.Height = PictureBox1.Height + Label1.Height + 40
    End Sub
End Class