Imports System.IO
Public Class Form1
    Public Sub ReadCSV()
        Using reader = New StreamReader("U:\test.csv")
            Dim displayName As List(Of String) = New List(Of String)()
            Dim email As List(Of String) = New List(Of String)()
            Dim emailAddress2 As List(Of String) = New List(Of String)()
            Dim emailAddress3 As List(Of String) = New List(Of String)()
            Dim homePhone As List(Of String) = New List(Of String)()
            Dim businessPhone As List(Of String) = New List(Of String)()
            Dim homeFax As List(Of String) = New List(Of String)()
            Dim pager As List(Of String) = New List(Of String)()
            Dim mobilePhone As List(Of String) = New List(Of String)()
            Dim homeStreet As List(Of String) = New List(Of String)()
            Dim homeCity As List(Of String) = New List(Of String)()
            Dim homeState As List(Of String) = New List(Of String)()
            Dim homeZip As List(Of String) = New List(Of String)()
            Dim busAddress As List(Of String) = New List(Of String)()
            Dim busCity As List(Of String) = New List(Of String)()
            Dim busState As List(Of String) = New List(Of String)()
            Dim busZip As List(Of String) = New List(Of String)()
            Dim jobTitle As List(Of String) = New List(Of String)()
            Dim organization As List(Of String) = New List(Of String)()
            Dim notes As List(Of String) = New List(Of String)()
            Dim birthday As List(Of String) = New List(Of String)()
            Dim webPage As List(Of String) = New List(Of String)()
            Dim webPage2 As List(Of String) = New List(Of String)()
            Dim categories As List(Of String) = New List(Of String)()

            While Not reader.EndOfStream
                Dim line = reader.ReadLine()
                Dim values = line.Split(","c)
                displayName.Add(values(0))
                email.Add(values(1))
                emailAddress2.Add(values(2))
                emailAddress3.Add(values(3))
                homePhone.Add(values(4))
                businessPhone.Add(values(5))
                homeFax.Add(values(6))
                pager.Add(values(7))
                mobilePhone.Add(values(8))
                homeStreet.Add(values(9))
                homeCity.Add(values(10))
                homeState.Add(values(11))
                homeZip.Add(values(12))
                busAddress.Add(values(13))
                busCity.Add(values(14))
                busState.Add(values(15))
                busZip.Add(values(16))
                jobTitle.Add(values(17))
                organization.Add(values(18))
                notes.Add(values(19))
                birthday.Add(values(20))
                webPage.Add(values(21))
                webPage2.Add(values(22))
                categories.Add(values(23))
            End While

            For i As Integer = 1 To displayName.Count()
                ListBox1.Items.Add(listA(i - 1))
                ListBox2.Items.Add(listB(i - 1))
            Next

        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReadCSV()
    End Sub
End Class
