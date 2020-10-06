Dim DirPath
DirPath=Environment.Value("ResultDir")
print(DirPath +DataTable("ResultDir", dtLocalSheet))

'Write into a file
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFile = FSO.CreateTextFile(DataTable("TxtFilePath", dtLocalSheet),True)
' Writes a specified string to the file
oFile.WriteLine(DirPath +DataTable("ResultDir", dtLocalSheet))
oFile.Close
