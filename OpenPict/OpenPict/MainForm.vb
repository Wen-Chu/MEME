Imports System.IO
Imports System.Data.SqlClient

Public Class MainForm
    Dim connection As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\VB\OpenPict\OpenPict\Picture.mdf;Integrated Security=True")
    Dim ssub As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\VB\OpenPict\OpenPict\Picture.mdf;Integrated Security=True")

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        SelectMode.Show()
        Me.Close()
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ssub As New SqlCommand("select * from Picture", connection)
        Dim table As New DataTable()
        Dim adapter As New SqlDataAdapter(ssub)
        Dim count As Int32 = adapter.Fill(table)

        Dim command As New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        Dim img() As Byte
        img = table.Rows(0)(1)
        Dim ms As New MemoryStream(img)
        PictureBox1.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 1
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox2.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 2
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox3.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 3
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox4.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 4
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox5.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 5
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox6.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 6
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox7.Image = Image.FromStream(ms)

        command = New SqlCommand("Select * from Picture where id = @id", connection)
        command.Parameters.Add("@id", SqlDbType.Int).Value = count - 7
        table = New DataTable()
        adapter = New SqlDataAdapter(command)

        adapter.Fill(table)
        img = table.Rows(0)(1)
        ms = New MemoryStream(img)
        PictureBox8.Image = Image.FromStream(ms)
    End Sub

    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreviewP1.Click, btnPreviewP2.Click, btnPreviewP3.Click, btnPreviewP4.Click, btnPreviewP5.Click, btnPreviewP6.Click, btnPreviewP7.Click, btnPreviewP8.Click
        Dim num1 As Integer
        num1 = DirectCast((sender), Button).Tag
        PreviewPic.PreviewPic_Show(num1)
        PreviewPic.Show()
    End Sub

    Private Sub btnDownload_Click(sender As Object, e As EventArgs) Handles btnDownload1.Click, btnDownload2.Click, btnDownload3.Click, btnDownload4.Click, btnDownload5.Click, btnDownload6.Click, btnDownload7.Click, btnDownload8.Click
        Try
            Dim myStream As Stream
            Dim saveFileDialog1 As New SaveFileDialog()

            saveFileDialog1.Filter = "Image (*.png)|*.png|All files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.RestoreDirectory = False

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                myStream = saveFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    ' Code to write the stream goes here.
                    Select Case DirectCast((sender), Button).Tag
                        Case 1
                            PictureBox1.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 2
                            PictureBox2.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 3
                            PictureBox3.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 4
                            PictureBox4.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 5
                            PictureBox5.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 6
                            PictureBox6.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 7
                            PictureBox7.Image.Save(saveFileDialog1.FileName + ".png")
                        Case 8
                            PictureBox8.Image.Save(saveFileDialog1.FileName + ".png")
                    End Select
                    myStream.Close()
                    My.Computer.FileSystem.DeleteFile(saveFileDialog1.FileName)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnDownload5.Click, btnDownload4.Click, btnDownload3.Click, btnDownload2.Click, btnDownload8.Click, btnDownload7.Click, btnDownload6.Click, btnDownload1.Click

    End Sub
End Class