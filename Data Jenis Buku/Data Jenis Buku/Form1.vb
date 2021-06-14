Imports System.Data.OleDb

Public Class DataJenis

    Sub Kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub

    Sub Isi()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilJenis()
        da = New OleDbDataAdapter("Select * From Jenis", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Jenis")
        DataGridView1.DataSource = ds.Tables("Jenis")
        DataGridView1.Refresh()
    End Sub
    Sub AturGrid()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 215
        DataGridView1.Columns(0).HeaderText = "KODE JENIS"
        DataGridView1.Columns(1).HeaderText = "NAMA JENIS"
    End Sub
    Private Sub DataJenis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilJenis()
        Call Kosong()
        Call AturGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Belum Lengkap!")
            TextBox1.Focus()
            Exit Sub
        Else
            cmd = New OleDbCommand("Select * from Jenis where KodeJenis='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim Simpan As String = "insert into Jenis(KodeJenis,Jenis) Values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                cmd = New OleDbCommand(Simpan, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("Simpan Data Sukses", MsgBoxStyle.Information, "Perhatian")
            End If
            ' Call Tampiljenis()
            ' Call Kosong()
            TextBox1.Focus()
        End If

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            TextBox2.Text = UCase(TextBox2.Text)
        End If
    End Sub


    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.TextBox2.Text = .Cells(1).Value
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Jenis Blum Diisi!")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim Ubah As String = "Update Jenis set" &
                "Jenis='" & TextBox2.Text & "' " &
                "where KodeJenis='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(Ubah, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("Ubah Data Sukses!", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call Kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Buku Blum Diisi!")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin Akan Menghapus Data Jenis" & TextBox1.Text & "?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                cmd = New OleDbCommand("Delete * From Jenis where KodeJenis='" & TextBox1.Text & "'", Conn)
                cmd.ExecuteNonQuery()
                Call Kosong()
                Call TampilJenis()
            Else
                Call Kosong()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call Kosong()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 2
        If e.KeyChar = Chr(13) Then
            cmd = New OleDbCommand("Selext * From Jenis where KodeJenis='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = True Then
                TextBox2.Text = rd.Item(1)
                TextBox2.Focus()
            Else
                Call Isi()
                TextBox2.Focus()
            End If
        End If
    End Sub


    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        cmd = New OleDbCommand("Select * From Jenis where KodeJenis like '%" & TextBox3.Text & "%'", Conn)
        rd = cmd.ExecuteReader
        rd.Read()
        If rd.HasRows Then
            da = New OleDbDataAdapter("Select * From Jenis where KodeJenis like'%" & TextBox3.Text & "%'", Conn)
            ds = New DataSet
            da.Fill(ds, "Dapat")
            DataGridView1.DataSource = ds.Tables("Dapat")
            DataGridView1.ReadOnly = True
        Else
            MsgBox("Data Tidak Ditemukan")
        End If
    End Sub
End Class
