Imports System.IO
Imports System.Text.RegularExpressions
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
    Dim firstName As List(Of String) = New List(Of String)()
    Dim lastName As List(Of String) = New List(Of String)()
    Dim partFirstName As List(Of String) = New List(Of String)()
    Dim partLastName As List(Of String) = New List(Of String)()
    Dim partFirstNameBlank As List(Of String) = New List(Of String)()
    Dim partLastNameBlank As List(Of String) = New List(Of String)()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadCSV()
    End Sub

    Public Sub ReadCSV()
        Using reader = New StreamReader("U:\mailinglistNEW.csv")

            Dim count As Integer = 0

            While Not reader.EndOfStream
                Dim line = reader.ReadLine()
                Dim values = line.Split("~"c)
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
                firstName.Add(values(24))
                lastName.Add(values(25))
                partFirstName.Add(values(26))
                partLastName.Add(values(27))
            End While

            RemoveEmptySpace(partFirstName, partFirstNameBlank)
            RemoveEmptySpace(partLastName, partLastNameBlank)

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
                ListBox25.Items.Add(firstName(i - 1))
                ListBox26.Items.Add(lastName(i - 1))
                ListBox27.Items.Add(partFirstName(i - 1))
                ListBox28.Items.Add(partLastName(i - 1))
                ListBox29.Items.Add(partFirstNameBlank.Count())
                ListBox30.Items.Add(partLastNameBlank.Count())
                count += 1
            Next

            lblCount.Text = count

            checkNames()

        End Using
    End Sub

    Public Sub checkNames()
        Dim count As Integer = 0

        For i As Integer = 1 To displayName.Count()
            If displayName.Item(i - 1).Contains("and") Then
                count += 1
            End If
        Next

        lblCount.Text = count

    End Sub

    Public Sub allWrite(ByVal selectedList As List(Of String))
        Threading.Thread.Sleep(2000)
        For i As Integer = 1 To selectedList.Count()
            SendKeys.Send(selectedList.Item(i - 1))
            SendKeys.Send("{ENTER}")
        Next
    End Sub

    Public Sub emailWrite(ByVal selectedList As List(Of String))
        Threading.Thread.Sleep(2000)
        For i As Integer = 1 To selectedList.Count()
            SendKeys.Send(selectedList.Item(i - 1).ToLower())
            SendKeys.Send("{ENTER}")
        Next
    End Sub

    Public Sub phoneWrite(ByVal selectedList As List(Of String))
        Threading.Thread.Sleep(2000)
        For i As Integer = 1 To selectedList.Count()
            Dim number As String = Regex.Replace(selectedList.Item(i - 1), "[^0-9]", "")

            If (number.Length = 11) Then
                number = number.Substring(1)
            End If
            SendKeys.Send(number)
            SendKeys.Send("{ENTER}")
        Next

    End Sub

    Public Sub RemoveEmptySpace(ByVal initialList As List(Of String), ByVal savedList As List(Of String))
        For i As Integer = 1 To initialList.Count()
            Dim line As String = initialList.Item(i - 1)

            If line Like "*[a-zA-Z]*" Then
                savedList.Add(line)
            End If
        Next
    End Sub

    Private Sub noFormatClick(sender As Object, e As EventArgs) Handles btnEmail.Click
        emailWrite(email)
    End Sub

    Private Sub btnHomeAddress_Click(sender As Object, e As EventArgs) Handles btnHomeAddress.Click
        allWrite(homeStreet)
    End Sub

    Private Sub btnHomeCity_Click(sender As Object, e As EventArgs) Handles btnHomeCity.Click
        allWrite(homeCity)
    End Sub

    Private Sub btnHomeState_Click(sender As Object, e As EventArgs) Handles btnHomeState.Click
        allWrite(homeState)
    End Sub

    Private Sub btnHomeZip_Click(sender As Object, e As EventArgs) Handles btnHomeZip.Click
        allWrite(homeZip)
    End Sub

    Private Sub btnMobile_Click(sender As Object, e As EventArgs) Handles btnMobile.Click
        phoneWrite(mobilePhone)
    End Sub

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click
        phoneWrite(homePhone)
    End Sub

    Private Sub btnHomeFax_Click(sender As Object, e As EventArgs) Handles btnHomeFax.Click
        phoneWrite(homeFax)
    End Sub

    Private Sub btnPager_Click(sender As Object, e As EventArgs) Handles btnPager.Click
        phoneWrite(pager)
    End Sub

    Private Sub btnBusiness_Click(sender As Object, e As EventArgs) Handles btnBusiness.Click
        phoneWrite(businessPhone)
    End Sub

    Private Sub btnBusinessAddress_Click(sender As Object, e As EventArgs) Handles btnBusinessAddress.Click
        allWrite(busAddress)
    End Sub

    Private Sub btnBusinessCity_Click(sender As Object, e As EventArgs) Handles btnBusinessCity.Click
        allWrite(busCity)
    End Sub

    Private Sub btnBusinessState_Click(sender As Object, e As EventArgs) Handles btnBusinessState.Click
        allWrite(busState)
    End Sub

    Private Sub btnBusinessZIP_Click(sender As Object, e As EventArgs) Handles btnBusinessZIP.Click
        allWrite(busZip)
    End Sub

    Private Sub btnJobTitle_Click(sender As Object, e As EventArgs) Handles btnJobTitle.Click
        allWrite(jobTitle)
    End Sub

    Private Sub btnOrganization_Click(sender As Object, e As EventArgs) Handles btnOrganization.Click
        allWrite(organization)
    End Sub

    Private Sub btnNotes_Click(sender As Object, e As EventArgs) Handles btnNotes.Click
        allWrite(notes)
    End Sub
    Private Sub btnCategories_Click(sender As Object, e As EventArgs) Handles btnCategories.Click
        allWrite(categories)
    End Sub

    Private Sub btnEmail2_Click(sender As Object, e As EventArgs) Handles btnEmail2.Click
        emailWrite(emailAddress2)
    End Sub

    Private Sub btnEmail3_Click(sender As Object, e As EventArgs) Handles btnEmail3.Click
        emailWrite(emailAddress3)
    End Sub

    Private Sub btnFirstName_Click(sender As Object, e As EventArgs) Handles btnFirstName.Click
        allWrite(firstName)
    End Sub

    Private Sub btnLastName_Click(sender As Object, e As EventArgs) Handles btnLastName.Click
        allWrite(lastName)
    End Sub

    Private Sub btnPartFirstName_Click(sender As Object, e As EventArgs) Handles btnPartFirstName.Click
        allWrite(partFirstName)
    End Sub

    Private Sub btnPartLastName_Click(sender As Object, e As EventArgs) Handles btnPartLastName.Click
        allWrite(partLastName)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnPartFirstBlank.Click
        allWrite(partFirstName)
    End Sub

    Private Sub btnPartLastBlank_Click(sender As Object, e As EventArgs) Handles btnPartLastBlank.Click
        allWrite(partLastName)
    End Sub
End Class
