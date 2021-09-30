Imports System.Data.OleDb
Imports System.Xml
Public Class Form1
    Dim cmd As New OleDbCommand
    Dim dt As New DataTable()
    Dim verial As New OleDbDataAdapter
    Dim bs As New BindingSource()
    Dim ds As New DataSet
    Dim okulno, ad, soyad, okul As String
    Dim conn As New OleDbConnection("Provider=Microsoft.Jet.Oledb.4.0;Data 
Source=C:\Users\kacan\Documents\ogrenci.mdb")
    Dim n, vize1, vize2, final, basari As Integer

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim dosya As New XmlDocument
        dosya.Load("C:\Users\kacan\source\repos\odev\ogr.xml")
        Dim liste As XmlNodeList = dosya.GetElementsByTagName("kisiler")
        Dim kisi As XmlNode
        If TextBox8.Text <> "" And IsNumeric(TextBox8.Text) Then
            For Each kisi In liste
                okul = kisi("okulno").FirstChild.Value
            Next
            If okul = TextBox8.Text Then
                MessageBox.Show("Bu okul no daha önce kayıt edilmiş!")
            Else
                Dim kisiler As XmlElement = dosya.CreateElement("kisiler")
                kisiler.SetAttribute("id", TextBox8.Text)
                Dim okulno As XmlNode = dosya.CreateNode("element", "okulno", "")
                okulno.InnerText = TextBox1.Text
                kisiler.AppendChild(okulno)
                Dim ad As XmlNode = dosya.CreateNode("element", "ad", "")
                ad.InnerText = TextBox2.Text
                kisiler.AppendChild(ad)
                Dim soyad As XmlNode = dosya.CreateNode("element", "soyad", "")
                soyad.InnerText = TextBox3.Text
                kisiler.AppendChild(soyad)
                Dim vize1 As XmlNode = dosya.CreateNode("element", "vize1", "")
                vize1.InnerText = TextBox4.Text
                kisiler.AppendChild(vize1)
                Dim vize2 As XmlNode = dosya.CreateNode("element", "vize2", "")
                vize2.InnerText = TextBox5.Text
                kisiler.AppendChild(vize2)
                Dim final As XmlNode = dosya.CreateNode("element", "final", "")
                final.InnerText = TextBox6.Text
                kisiler.AppendChild(final)
                Dim basarinotu As XmlNode = dosya.CreateNode("element", "basarinotu", "")
                basarinotu.InnerText = TextBox7.Text
                kisiler.AppendChild(basarinotu)
                dosya.DocumentElement.AppendChild(kisiler)
                dosya.Save("C:\Users\kacan\source\repos\odev\ogr.xml")
                MessageBox.Show("Xml dosyasına kayıt başarıyla yapıldı")
            End If
        Else
            MessageBox.Show("Id girişi atlanamaz!")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        conn.Open()
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "delete from ders where okul_no='" & TextBox1.Text & "'"
        conn.Close()
        MessageBox.Show("Kayıt silindi")
        yenile()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "update ders set okul_no='" & TextBox1.Text & "',adi='" & TextBox2.Text &
        "',soyadi='" & TextBox3.Text & "',vize1='" & TextBox4.Text & "',vize2='" & TextBox5.Text & "',final='" &
        TextBox6.Text & "'"
        cmd.ExecuteNonQuery()
        conn.Close()
        MessageBox.Show("Kayıtlar güncellendi")
        yenile()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        n = Convert.ToInt32(InputBox("Kaç kişi gireceksiniz?"))
        For i As Integer = 1 To n
            okulno = InputBox("Okul no giriniz?")
            ad = InputBox("Ad giriniz?")
            soyad = InputBox("Soyad giriniz?")
            vize1 = Convert.ToInt32(InputBox("1. vize notunu giriniz?"))
            vize2 = Convert.ToInt32(InputBox("2. vize notunu?"))
            final = Convert.ToInt32(InputBox("Final notunu giriniz?"))
            basari = b.hesapla(vize1, vize2, final)
        Next
        If basari > 0 Then
            MessageBox.Show("Başarı notu hesaplama işlemi yapıldı")
        End If
        conn.Open()
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "insert into ders(okul_no,adi,soyadi,vize1,vize2,final,basari_notu) values('" & okulno &
        "','" & ad & "','" & soyad & "','" & vize1 & "','" & vize2 & "','" & final & "','" & basari & "')"
        cmd.ExecuteNonQuery()
        conn.Close()
        MessageBox.Show("Öğrenciler veri tabanına aktarıldı")
        yenile()
    End Sub

    Private ReadOnly Property InputBox(v As String) As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Dim b As New hesap.Class1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ReadOnly = True
        TextBox7.Enabled = False
        conn.Open()
        verial = New OleDbDataAdapter("select * from ders", conn)
        verial.Fill(dt)
        bs.DataSource = dt
        DataGridView1.DataSource = bs
        TextBox1.DataBindings.Add("Text", bs, "okul_no")
        TextBox2.DataBindings.Add("Text", bs, "adi")
        TextBox3.DataBindings.Add("Text", bs, "soyadi")
        TextBox4.DataBindings.Add("Text", bs, "vize1")
        TextBox5.DataBindings.Add("Text", bs, "vize2")
        TextBox6.DataBindings.Add("Text", bs, "final")
        TextBox7.DataBindings.Add("Text", bs, "basari_notu")
        conn.Close()
    End Sub
    Sub yenile()
        Dim verial As New OleDbDataAdapter("select * from ders", conn)
        Dim dt As New DataTable()
        verial.Fill(dt)
        BindingSource1.DataSource = dt
        DataGridView1.DataSource = BindingSource1
    End Sub
End Class
