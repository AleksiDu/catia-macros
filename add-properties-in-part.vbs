Sub CATMain()

Set partDocument1 = CATIA.ActiveDocument
Set product1 = partDocument1.GetItem("")

On Error Resume Next
Err.Clear
Set strParam1 = product1.UserRefProperties.Item("Stock_Size")
if Err.Number = 0 then
    'Do Nothing
else
Set parameters1 = product1.UserRefProperties
Set strParam1 = parameters1.CreateString("Stock_Size", "")
strParam1.ValuateFromString ""
end if

On Error Resume Next
Err.Clear
Set strParam2 = product1.UserRefProperties.Item("Material_")
if Err.Number = 0 then
    'Do Nothing
else
Set parameters2 = product1.UserRefProperties
Set strParam2 = parameters2.CreateString("Material_", "")
strParam2.ValuateFromString ""
end if
End Sub