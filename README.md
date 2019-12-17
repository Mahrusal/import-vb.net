Imports System.Data.SqlClient
Public Class Form1
    Private Function saveData(sql As String)
        'Dim sqlCOn As SqlConnection = New SqlConnection("Data Source=localhost;Initial Catalog=userad;Integrated Security=True")
        Dim sqlCOn As SqlConnection = New SqlConnection("Data Source=localhost;Initial Catalog=OLAH_AXIST;Integrated Security=True")
        Dim sqlCmd As SqlCommand
        Dim resul As Boolean

        Try

            sqlCOn.Open()
            sqlCmd = New SqlCommand
            With sqlCmd
                .Connection = sqlCOn
                .CommandText = sql
                resul = .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlCOn.Close()
        End Try
        Return resul
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Filter = "Excel files(*.xlsx)|*.xlsx|All files (*.*)|*.*"
            .FilterIndex = 1
            .Title = "Import data from Excel file"
        End With
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            MsgBox("Success")
            ProgressBar1.Value = 0
        Else
            ProgressBar1.Value += 1
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim OLEcon As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & TextBox1.Text & " ; " & "Extended Properties=Excel 8.0;")
        Dim OLEcmd As New OleDb.OleDbCommand
        Dim OLEda As New OleDb.OleDbDataAdapter
        Dim OLEdt As New DataTable
        Dim sql As String
        Dim resul As Boolean

        Try
            OLEcon.Open()
            With OLEcmd
                .Connection = OLEcon
                .CommandText = "select * from [Sheet1$]"
            End With
            OLEda.SelectCommand = OLEcmd
            OLEda.Fill(OLEdt)

            For Each r As DataRow In OLEdt.Rows
                'sql = "INSERT INTO tbpegawai (IDPegawai ,Nama,Nip,UnitJenjang) VALUES ('" & r(0).ToString & "','" & r(1).ToString & "','" & r(2).ToString & "','" & r(3).ToString & "')"
                sql = "INSERT INTO AXTO (prospect_id,name,campaign_id,product_id,custname,dob,mphone,mphone2,call_id,calldate,premium,anp,tso_id,spv_id,card_type,cifnumber,accnumber,anp1,tso_id_1,policy_id,acctnum,payment_methode,Refferenc_product) VALUES ('" & r(0).ToString & "','" & r(1).ToString & "','" & r(2).ToString & "','" & r(3).ToString & "','" & r(4).ToString & "','" & r(5).ToString & "','" & r(6).ToString & "','" & r(7).ToString & "','" & r(8).ToString & "','" & r(9).ToString & "','" & r(10).ToString & "','" & r(11).ToString & "','" & r(12).ToString & "','" & r(13).ToString & "','" & r(14).ToString & "','" & r(15).ToString & "','" & r(16).ToString & "','" & r(17).ToString & "','" & r(18).ToString & "','" & r(19).ToString & "','" & r(20).ToString & "','" & r(21).ToString & "','" & r(22).ToString & "')"
                resul = saveData(sql)
                If resul Then
                    Timer1.Start()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            OLEcon.Close()
        End Try
    End Sub
End Class
