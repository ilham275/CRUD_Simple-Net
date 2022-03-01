Imports System.Data.Odbc
Public Class Form1
    Dim Coon As OdbcConnection
    Dim Cmd As OdbcCommand
    Dim Ds As DataSet
    Dim Da As OdbcDataAdapter
    Dim Rd As OdbcDataReader
    Dim MyDB As String
    Sub Koneksi()
        MyDB = "DSN=db_mahasiswa"
        Coon = New OdbcConnection(MyDB)
        If Coon.State = ConnectionState.Closed Then Coon.Open()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        Button1.Text = "INPUT"
        Button2.Text = "EDIT"
        Button3.Text = "DELETE"
        Button4.Text = "BATAL"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Call Koneksi()
        Da = New OdbcDataAdapter("Select * From tbl_Dsiswa", Coon)
        Ds = New DataSet
        Da.Fill(Ds, "tbl_Dsiswa")
        DataGridView1.DataSource = Ds.Tables("tbl_Dsiswa")
    End Sub
    Sub FieldAktif()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox1.Focus()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Button1.Text = "SIMPAN"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call FieldAktif()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("pastikan semua field terisi")
            Else
                Call Koneksi()
                Dim InputData As String = "Insert into tbl_Dsiswa values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                Cmd = New OdbcCommand(InputData, Coon)
                Cmd.ExecuteNonQuery()
                MsgBox("Input Data Berhasil")
                Call KondisiAwal()

            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Button2.Text = "UPDATE"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call FieldAktif()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("pastikan semua field terisi")
            Else
                Call Koneksi()
            Dim EditData As String = "Update tbl_Dsiswa set nama='" & TextBox2.Text & "', telp='" & TextBox3.Text & "',alamat='" & TextBox4.Text & "' where nim='" & TextBox1.Text & "'"
            Cmd = New OdbcCommand(EditData, Coon)
            Cmd.ExecuteNonQuery()
            MsgBox("Edit Data Berhasil")
            Call KondisiAwal()

        End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New OdbcCommand("Select * from tbl_Dsiswa where nim ='" & TextBox1.Text & "'", Coon)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("nama")
                TextBox3.Text = Rd.Item("telp")
                TextBox4.Text = Rd.Item("alamat")

            Else
                MsgBox("Data tidak ada")
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "DELETE" Then
            Button3.Text = "HAPUS"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "BATAL"
            Call FieldAktif()
        Else

            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("pastikan semua field terisi")
            Else
                Call Koneksi()
                Dim deleteData As String = "delete from tbl_Dsiswa where nim='" & TextBox1.Text & "'"
                Cmd = New OdbcCommand(deleteData, Coon)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di hapus")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button4.Text = "TUTUP" Then
            End
        Else
            Call KondisiAwal()
        End If
    End Sub
End Class
