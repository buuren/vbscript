Const olContactItem = 2a
Dim objOutl, objContact
Set objOutl = WScript.CreateObject("Outlook.Application")
Set objContact = objOutl.CreateItem(olContactItem)
objContact.FirstName = "john"
objContact.LastName = "smith"
objContact.Email1Address = "john@world.com"
objContact.Save()