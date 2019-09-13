Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory
Set WshShell = Nothing
Set FSO = CreateObject("Scripting.FileSystemObject")

Set Folder = FSO.GetFolder(strCurDir)
Set fls=Folder.Files
if Wscript.Arguments.Count<2 then
  rfrom=InputBox("Что искать","Поиск замена в XLS")
  rto=InputBox("На что заменить","Поиск замена в XLS")  
else
  rfrom=Wscript.Arguments(0)
  rto=Wscript.Arguments(1)
end if
'Wscript.Echo(rfrom+":"+rto)
Set xlobj=CreateObject("Excel.Application")
xlobj.DisplayAlerts=false
for each File in folder.files 
  fname= File.Name
  if Right(fname,4)=".xls" then 
    Work(File.path)
  end if
  if Right(fname,5)=".xlsx" then
    Work(File.path)
  end if
 
next 
xlobj.Quit

function Work(fname)
  xlobj.workbooks.Open fname
  for i=1 to xlobj.ActiveWorkbook.Sheets.Count
  xlobj.ActiveWorkBook.Sheets(i).Cells.Replace rfrom, rto
  next
  xlobj.ActiveWorkbook.Close true

end function