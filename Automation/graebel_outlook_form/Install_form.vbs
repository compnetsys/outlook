Const olPersonalRegistry = 2
Set objOL = CreateObject("Outlook.Application")
Set objItem = objOL.CreateItemFromTemplate("formlocation")
Set objFD = objItem.FormDescription
objFD.Name = "Appointment"
objFD.DisplayName = "Appointments"
objFD.PublishForm olPersonalRegistry
 
