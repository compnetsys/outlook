Const olPersonalRegistry = 2
Set objOL = CreateObject("Outlook.Application")
Set objItem = objOL.CreateItemFromTemplate("C:\github\outlook\Automation\Appointments.oft")
Set objFD = objItem.FormDescription
objFD.Name = "Appointment"
objFD.DisplayName = "Appointments"
objFD.PublishForm olPersonalRegistry
 
