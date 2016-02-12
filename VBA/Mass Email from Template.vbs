Sub Amsterdam()
    Dim email_arr
    email_arr = Array("mona@datasciencedojo.com", "janella@datasciencedojo.com", "brian@datasciencedojo.com", "raja@datasciencedojo.com", "kaitlyn@datasciencedojo.com", "beijuan@datasciencedojo.com", "phuc@datasciencedojo.com", "chenoa@datasciencedojo.com")

    Dim first_name_arr
    first_name_arr = Array("Mona", "Janella", "Brian", "Raja", "Kaitlyn", "Beijuan", "Phuc", "Chenoa")

    Dim myOlApp As Outlook.Application
    Dim MyItem As Outlook.MailItem
    Set myOlApp = CreateObject("Outlook.Application")

    For i = LBound(email_arr) To UBound(email_arr)
        Set MyItem = myOlApp.CreateItemFromTemplate("C:\gitrepos\non-commit\Amsterdam  Data Science  Engineering Bootcamp.oft")
        MyItem.HTMLBody = Replace(MyItem.HTMLBody, "qwerasdf", first_name_arr(i))
        MyItem.Recipients.Add (email_arr(i))
        MyItem.Send
    Next i
End Sub