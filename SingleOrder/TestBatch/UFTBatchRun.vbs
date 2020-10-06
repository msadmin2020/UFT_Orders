wScript.sleep 10000 

Set qtpApp=CreateObject("QuickTest.Application")
Set fsObj=CreateObject("Scripting.FileSystemObject")
Set qtpResObj=CreateObject("QuickTest.RunResultsOptions")

sFolderPath="E:\UFT_GIT\Git_UFTOrders\UFT_Orders"

qtpApp.Launch
qtpApp.Visible=True

Set mainFolderObj= fsObj.GetFolder(sFolderPath)
Set testSubFolders=mainFolderObj.SubFolders

For each folderObj in testSubFolders
     chkFolderObj = folderObj.Path & "\Action0"
     If (fsObj.FolderExists(chkFolderObj)) Then 'The folder is a QTP test folder
         
         qtpApp.Open folderObj.Path, True

          set qtTest=qtpApp.Test

          qtTest.Run

          strResult=qtpApp.Test.LastRunResults.Status
          'wScript.echo strResult


          wScript.sleep 10000 

          qtTest.close

          
      End If
Next

Set qtTest=nothing
Set qtpApp=nothing
Set fsObj=nothing 
Set qtpResObj=nothing