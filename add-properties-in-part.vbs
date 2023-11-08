Sub CATMain()

Set partDocument1 = CATIA.ActiveDocument
Set product1 = partDocument1.GetItem("")

Set parameters1 = product1.UserRefProperties
Set strParam1 = parameters1.CreateString("Stock_Size", "")
strParam1.ValuateFromString ""

Set parameters2 = product1.UserRefProperties
Set strParam2 = parameters2.CreateString("Stock_Size", "")
strParam2.ValuateFromString ""

End Sub