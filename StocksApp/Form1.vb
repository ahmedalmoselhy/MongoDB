Imports System
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Drawing.Image
Imports MongoDB.Bson
Imports MongoDB.Bson.IO
Imports MongoDB.Bson.Serialization
Imports MongoDB.Driver
Imports MongoDB.Driver.Linq
Imports MongoDB.Driver.Builders
Public Class Form1
    Dim client As MongoClient
    Dim db As IMongoDatabase
    Dim collection As IMongoCollection(Of BsonDocument)
    Dim dt As New DataTable
    Dim flag As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' connect to the database
        client = New MongoClient("mongodb://localhost:27017/")
        db = client.GetDatabase("db")
        collection = db.GetCollection(Of BsonDocument)("Models")
        dt.Columns.Add("id", Type.GetType("System.String"))
        dt.Columns.Add("Model Type", Type.GetType("System.String"))
        dt.Columns.Add("Image", Type.GetType("System.Byte[]"))
    End Sub
    ' add elements to the database
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ms As MemoryStream = New MemoryStream()
        PictureBox1.BackgroundImage.Save(ms, ImageFormat.Png)
        Dim arr As Byte() = ms.ToArray()
        Dim Models As BsonDocument = New BsonDocument
        With Models
            .Add("_id", TextBox1.Text)
            .Add("_ModelType", TextBox2.Text)
            .Add("_Image", arr)
        End With
        Try
            collection.InsertOne(Models)
        Catch ex As Exception
            MsgBox("id is unique, choose a different one !")
        End Try
    End Sub
    ' view all records in datagrid view
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, Button4.Click, Button5.Click
        If flag = True Then
            DataGridView1.DataSource.clear()
        End If
        Dim filter = Builders(Of BsonDocument).Filter.Empty
        For Each item As BsonDocument In collection.Find(filter).ToList
            Dim id As BsonElement = item.GetElement("_id")
            Dim ModelType As BsonElement = item.GetElement("_ModelType")
            Dim image As BsonElement = item.GetElement("_Image")
            Dim arr As Byte() = image.Value
            Dim ms As MemoryStream = New MemoryStream(arr)
            dt.Rows.Add(id.Value, ModelType.Value, ms.ToArray())
        Next
        DataGridView1.DataSource = dt
        flag = True
    End Sub
    ' choose an image
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Image Files |*.bmp;*.gif;*.jpg;*.png;*.tif|Allfiles|*.*"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim ConvertedImg = New Bitmap(Image.FromFile(OpenFileDialog1.FileName, True), 120, 120)
            PictureBox1.BackgroundImage = ConvertedImg
        End If
    End Sub
    'exit
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        System.Environment.Exit(0)
    End Sub
    ' search for a specific record
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim filter = Builders(Of BsonDocument).Filter.Empty
        For Each item As BsonDocument In collection.Find(filter).ToList
            Dim id As BsonElement = item.GetElement("_id")
            Dim ModelType As BsonElement = item.GetElement("_ModelType")
            Dim image As BsonElement = item.GetElement("_Image")
            Dim arr As Byte() = image.Value
            Dim ms As MemoryStream = New MemoryStream(arr)
            If TextBox1.Text = id.Value Then
                TextBox2.Text = ModelType.Value
                PictureBox1.BackgroundImage = FromStream(ms)
                Exit Sub
            End If
        Next
        MsgBox("No Models found for this id !")
        TextBox1.Clear()
        TextBox2.Clear()
        PictureBox1.BackgroundImage = My.Resources._Default
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString
        TextBox2.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString
        Dim arr As Byte() = DataGridView1.SelectedRows(0).Cells(2).Value
        Dim ms As MemoryStream = New MemoryStream(arr)
        PictureBox1.BackgroundImage = FromStream(ms)
    End Sub
    ' remove a record
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim q = Builders(Of BsonDocument).Filter.Eq(Of String)("_id", DataGridView1.SelectedRows(0).Cells(0).Value.ToString)
        collection.DeleteOne(q)
        TextBox1.Clear()
        TextBox2.Clear()
        PictureBox1.BackgroundImage = My.Resources._Default
        Button3.PerformClick()
    End Sub
    ' edit a record
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, Button8.Click
        TextBox1.Enabled = False
        Button2.Enabled = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = True
    End Sub
    ' save edited record
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of String)("_id", TextBox1.Text)
        Dim ms As MemoryStream = New MemoryStream()
        PictureBox1.BackgroundImage.Save(ms, ImageFormat.Png)
        Dim arr As Byte() = ms.ToArray()
        Dim Models As BsonDocument = New BsonDocument
        With Models
            .Add("_id", TextBox1.Text)
            .Add("_ModelType", TextBox2.Text)
            .Add("_Image", arr)
        End With
        collection.ReplaceOne(fltr, Models)
        TextBox1.Enabled = True
        Button2.Enabled = True
        Button5.Enabled = True
        Button3.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = False
        Button3.PerformClick()
    End Sub
End Class