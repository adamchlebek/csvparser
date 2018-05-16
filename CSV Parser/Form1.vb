Imports System.IO
Public Class Form1

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


    Public Sub ReadCSV()
        Using reader = New StreamReader("U:\mailinglist.csv")


            Dim count As Integer = 0

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
                ListBox1.Items.Add(displayName(i - 1))
                ListBox2.Items.Add(email(i - 1))
                ListBox3.Items.Add(emailAddress2(i - 1))
                ListBox4.Items.Add(emailAddress3(i - 1))
                ListBox5.Items.Add(homePhone(i - 1))
                ListBox6.Items.Add(businessPhone(i - 1))
                ListBox7.Items.Add(homeFax(i - 1))
                ListBox8.Items.Add(pager(i - 1))
                ListBox9.Items.Add(mobilePhone(i - 1))
                ListBox10.Items.Add(homeStreet(i - 1))
                ListBox11.Items.Add(homeCity(i - 1))
                ListBox12.Items.Add(homeState(i - 1))
                ListBox13.Items.Add(homeZip(i - 1))
                ListBox14.Items.Add(busAddress(i - 1))
                ListBox15.Items.Add(busCity(i - 1))
                ListBox16.Items.Add(busState(i - 1))
                ListBox17.Items.Add(busZip(i - 1))
                ListBox18.Items.Add(jobTitle(i - 1))
                ListBox19.Items.Add(organization(i - 1))
                ListBox20.Items.Add(notes(i - 1))
                ListBox21.Items.Add(birthday(i - 1))
                ListBox22.Items.Add(webPage(i - 1))
                ListBox23.Items.Add(webPage2(i - 1))
                ListBox24.Items.Add(categories(i - 1))
                count += 1
            Next

            Button2.Text = count

            checkNames()

        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReadCSV()
    End Sub

    Public Sub checkNames()
        Dim count As Integer = 0

        For i As Integer = 1 To displayName.Count()
            If displayName.Item(i - 1).Contains("and") Then
                count += 1
            End If
        Next

        Button2.Text = count

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        emailWrite(email)
    End Sub

    Public Sub emailWrite(ByVal selectedList As List(Of String))
        Dim count As Integer = 0
        Threading.Thread.Sleep(2000)
        For i As Integer = 1 To selectedList.Count()
            SendKeys.Send(selectedList.Item(i - 1))
            SendKeys.Send("{ENTER}")
        Next

        Button3.Text = count
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadCSV()
    End Sub
End Class
