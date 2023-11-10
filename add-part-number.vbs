Sub CATMain()

Set partDocument1 = CATIA.ActiveDocument
Set part1 = partDocument1.Part
Set parameters1 = part1.Parameters
Set partDocument1 = CATIA.ActiveDocument

Set product1 = partDocument1.GetItem("")

On Error Resume Next
Err.Clear
Set strParam1 = parameters1.Item("Search_string")
if Err.Number = 0 then
    'Nothing TODO if Parameter Exists
else
    'TODO Stuff if paramete dose not Exist
    Set parameters1 = part1.Parameters
    Set strParam1 = parameters1.CreateString("", "")
    strParam1.Rename "Search_string"
    part1.Update
    strParam1.Value = "-"
 end if

Set parameters2 = part1.Parameters
On Error Resume Next
Err.Clear
Set intParam1 = parameters2.Item("Search-result")
if Err.Number = 0 then
    'Nothing
else
Set intParam1 = parameters2.CreateInteger("",0)
intParam1.Rename "Search_result"
Set relations1 = part1.Relations
Set product1 = partDocument1.product
Set formula1 = relations1.CreateFormula("Formula.1", "", intParam1, "`" & product1.PartNumber & "\Part Number` .Search(Search_string)")
formula1.Rename = "Formula.1"
end if

Set parameters3 = part1.Parameters
On Error Resume Next
Err.Clear
Set strParam2 = parameters3.Item("Code")
if Err.Number = 0 then
    'Nothing
else
Set strParam2 = parameters3.CreateString("","")
strParam2.Rename "Code"
part1.Update
Set product2 = partDocument1.product
Set relations2 = part1.Relations
formulaStr = "`" & product1.PartNumber & "\Part Number` -> Extract(0,Search_result)"
Set formula2 = relations2.CreateFormula("Formula.2", "", strParam2, formulaStr)
part1.Update
end if


On Error Resume Next
Err.Clear
Set strParam3 = parameters3.Item("Code")
if strParam3.value = "MF" then
    product1.Revision = 1
    product1.Definition = parameters1.Item("Code").value
    product1.DescriptionRef = "Milling Fixture"
end if
if  strParam3.value = "CL" then
    product1.Revision = 1
    product1.Definition = parameters1.Item("Code").value
    product1.DescriptionRef = "Clamp"
end if

On Error Resume Next
Err.Clear
Set strParam4 = product1.UserRefProperties.Item("Material_")
if Err.Number = 0 then
    strParam4.ValuateFromString "AL6061-T651"
else
    Set parameters3 = product1.UserRefProperties
    Set strParam3 = parameters3.CreateString("Material_", "")
    Set strParam4 = parameters3.CreateString("Stock_Size", "")
    strParam3.value = "AL6061-T651"
    strParam4.ValuateFromString ""
end if

End Sub