Public Class Form1
    Public isl As Boolean = False
    Dim stepee As Integer = 0
    Dim ref_i As Integer = 5
    Dim usid As String
    Dim forumc As String
    Sub ck(ByVal z As String)
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("value") = z Then
                ele.InvokeMember("click")
            End If
        Next
    End Sub
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        wb1.Navigate(lb1.SelectedItem.replace("act=ST&f", "act=Post&CODE=08&f"))
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("value") = "Submit Modified Post" Then
                ele.InvokeMember("click")
            End If
        Next
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("value") = "Go" Then
                ele.InvokeMember("click")
            End If
        Next
    End Sub

    Private Sub wb1_ProgressChanged(sender As Object, e As WebBrowserProgressChangedEventArgs)
        On Error Resume Next
        tsp.Maximum = e.MaximumProgress
        tsp.Value = e.CurrentProgress
    End Sub

    Private Sub tm_Tick(sender As Object, e As EventArgs) Handles tm.Tick
        If stepee = 1 Then ToolStripButton3.PerformClick()
        If stepee = 2 Then
            ToolStripButton3.PerformClick()
        End If
        If stepee = 3 Then
            ToolStripButton3.PerformClick()
            ToolStripButton4.PerformClick()
        End If
        If stepee = 4 Then
            ToolStripButton4.PerformClick()
            If Not lb1.SelectedIndex = lb1.Items.Count - 1 Then
                lb1.SelectedIndex += 1
            Else
                tm.Enabled = False
                Exit Sub
            End If
        End If
        tm.Enabled = False
        ToolStripButton7.PerformClick()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("name") = "TopicTitle" Then
                ele.InnerText = "------"
            End If
            If ele.GetAttribute("name") = "TopicDesc" Then
                ele.InnerText = "------"
            End If
            If ele.GetAttribute("name") = "Post" Then
                ele.InnerText = "This post has been overwritten by a script."
            End If
        Next
    End Sub


    Private Sub ToolStripButton7_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        If Not stepee >= 4 Then
            stepee += 1
        Else
            stepee = 1
        End If
        tm.Enabled = True
    End Sub
    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        For Each ClientControl As HtmlElement In wb1.Document.Links
            If ClientControl.GetAttribute("href").Contains("https://forum.lowyat.net/index.php?act=ST&f=" & forumc) Then
                lb1.Items.Add(ClientControl.GetAttribute("href").Replace("act=ST&f", "act=Post&CODE=08&f"))
            End If
        Next
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("title") = "Next page" Then
                ele.InvokeMember("click")
            End If
        Next
    End Sub

    Private Sub lb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb1.SelectedIndexChanged
        tm.Enabled = False
        wb1.Navigate(lb1.SelectedItem)
        stepee = 0
        tm2.Enabled = True
        tsp.Maximum = lb1.Items.Count - 1
        tsp.Value = lb1.SelectedIndex
    End Sub

    Private Sub tm2_Tick(sender As Object, e As EventArgs) Handles tm2.Tick
        ToolStripButton7.PerformClick()
        tm2.Enabled = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wb1.Navigate("https://forum.lowyat.net/index.php?act=Login&CODE=00")
    End Sub

    Private Sub tm_ref_Tick(sender As Object, e As EventArgs) Handles tm_ref.Tick
        If ref_i = 0 Then
            l_ref.PerformClick()
        Else
            ref_i -= 1
            l_ref.Text = "Refreshing in " & ref_i & " seconds"
        End If
    End Sub

    Private Sub l_ref_Click(sender As Object, e As EventArgs) Handles l_ref.Click
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("href").Contains("https://forum.lowyat.net/index.php?act=Login&CODE=03&id=") Then
                usid = ele.GetAttribute("href").Replace("https://forum.lowyat.net/index.php?act=Login&CODE=03&id=", "")
                stat.Text = "User ID: " & usid
            End If
            If ele.GetAttribute("id") = "userlinks" Then
                tm_ref.Enabled = False
                l_ref.Text = "Manual Refresh"
            Else
                ref_i = 5
            End If
        Next
    End Sub

    Private Sub ForumMainPageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForumMainPageToolStripMenuItem.Click
        wb1.Navigate("http://forum.lowyat.net")
    End Sub

    Private Sub SeeUsersPostToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeeUsersPostToolStripMenuItem.Click
        wb1.Navigate("https://forum.lowyat.net/index.php?act=Search&CODE=getalluser&mid=" & usid)

    End Sub

    Private Sub SeeUsersTopicsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeeUsersTopicsToolStripMenuItem.Click
        wb1.Navigate("https://forum.lowyat.net/index.php?act=Search&CODE=gettopicsuser&mid=" & usid)
    End Sub

    Private Sub UserProfileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserProfileToolStripMenuItem.Click
        wb1.Navigate("https://forum.lowyat.net/index.php?showuser=" & usid)
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

    End Sub

    Private Sub cust_forumid_TextChanged(sender As Object, e As EventArgs) Handles cust_forumid.TextChanged
        forumc = cust_forumid.Text
    End Sub

    Private Sub KopitiamToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KopitiamToolStripMenuItem.Click
        forumc = "23"
    End Sub

    Private Sub AllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllToolStripMenuItem.Click
        forumc = ""
    End Sub

    Private Sub tm3_Tick(sender As Object, e As EventArgs) Handles tm3.Tick
        For Each ele As HtmlElement In wb1.Document.All
            If ele.GetAttribute("title") = "Next page" Then
                ToolStripButton9.PerformClick()
                Exit Sub
            End If
        Next
        ToolStripButton6.Checked = False
        tm3.Enabled = False
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        tm3.Enabled = ToolStripButton6.Checked
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        lb1.SelectedIndex = 0
        pst.Text = "Purge in progress..."
    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        tm.Enabled = False
        tm2.Enabled = False
        pst.Text = "On hold"
    End Sub
End Class
