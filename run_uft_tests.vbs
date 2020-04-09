'======================== Functions ========================
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Function getFolder(folderName) 
    Dim fsObject, folder
    Set fsObject = CreateObject("Scripting.FileSystemObject")
    Set folder = fsObject.GetFolder(folderName)

    getFolder = folder
End Function

Function listSubFolders(folderName) 
    Dim files, folder, fsObject
    Set fsObject = CreateObject("Scripting.FileSystemObject")
    Set folder = fsObject.GetFolder(folderName)
    Set files = folder.subFolders

    Set listSubFolders = files
End Function


'======================== Main Script ========================
Dim args, baseTestPath, webAppUrl, resultBasePath, browser, failOnError, failOnWarning, numArgs, files, QTP, t
Set args = WScript.Arguments
baseTestPath = args.Item(0)
webAppUrl = null
resultBasePath = null
browser = null
failOnError = false
failOnWarning = false

numArgs = (args.Count - 1) / 2

For i = 0 to (numArgs - 1)
  If StrComp(args.Item(1 + (i*2)), "-w") = 0 Then
     webAppUrl = args.Item(1 + (i*2) + 1)
  End If
  If StrComp(args.Item(1 + (i*2)), "-r") = 0 Then
     resultBasePath = args.Item(1 + (i*2) + 1)
  End If
  If StrComp(args.Item(1 + (i*2)), "-b") = 0 Then
     browser = args.Item(1 + (i*2) + 1)
  End If
  If StrComp(args.Item(1 + (i*2)), "-e") = 0 Then
     failOnError = args.Item(1 + (i*2) + 1)
  End If
  If StrComp(args.Item(1 + (i*2)), "-f") = 0 Then
     failOnWarning = args.Item(1 + (i*2) + 1)
  End If
Next

WScript.StdOut.WriteLine "baseTestPath : " & baseTestPath
WScript.StdOut.WriteLine "webAppUrl : " & webAppUrl
WScript.StdOut.WriteLine "resultBasePath : " & resultBasePath
WScript.StdOut.WriteLine "browser : " & browser
WScript.StdOut.WriteLine "failOnError : " & failOnError
WScript.StdOut.WriteLine "failOnWarning : " & failOnWarning
If Not fso.FolderExists(baseTestPath) Then
   WScript.StdOut.WriteLine "Base Test Path does not exist!"
   WScript.Quit 1
End If

Set files = listSubFolders(baseTestPath)

Set QTP = CreateObject("QuickTest.Application")
t = true

For Each file in files
    Dim testPath
    testPath = baseTestPath & "\" & file.Name
    WScript.StdOut.WriteLine "Opening Test : " & testPath
    QTP.Open testPath, true
    Dim test
    Set test = QTP.Test
    
    'set webAppUrl
    If Not isNull(webAppUrl) Then
        Dim settings, launchers, webLauncher
        Set settings = test.Settings
        Set launchers = settings.Launchers
        Set webLauncher = launchers.Item("Web")
        webLauncher.Address = webAppUrl
        WScript.StdOut.WriteLine "Setting App Url on Settings.Launchers(Web).Address"
    End If
   
    ' set browser
    If Not isNull(browser) Then
        Dim asettings, alaunchers, awebLauncher
        Set asettings = test.Settings
        Set alaunchers = asettings.Launchers
        Set awebLauncher = alaunchers.Item("Web")
        awebLauncher.Browser = browser
        WScript.StdOut.WriteLine "Setting Browser on Settings.Launchers(Web).Browser"
    End If

    'redirect output
    If isNull(resultBasePath) Then
        WScript.stdOut.WriteLine "Executing test..."
        test.Run
    Else
        Dim qtpResult, testResult

        If Not fso.FolderExists(resultBasePath) Then
            WScript.StdOut.WriteLine "Result base path does not exist!"
            WScript.Quit 1
        End If

        testResult = resultBasePath & file.Name
 
        Set qtpResult = CreateObject("QuickTest.RunResultsOptions")
        qtpResult.ResultsLocation = testResult
        WScript.stdOut.WriteLine "Executing test..."
        test.Run qtpResult
    End If

    Dim isRunning 
    isRunning = true
    
    'sleep util done
    While isRunning
        WScript.sleep(2000)
        isRunning = test.IsRunning
    WEnd
        
    'check for failures
    If failOnError Or failOnWarning Then
        Dim results, status
        Set results = test.LastRunResults
        status = results.Status
        
        If status = "Failed" And failOnError Then
             WScript.stdOut.WriteLine "QTP Test Failed!!!"
             WScript.Quit 1
        End If
 
        If status = "Warning" And failOnWarning Then
             WScript.stdOut.WriteLine "QTP Test had a warning!!!"
             WScript.Quit 1
        End If
    End If
Next

WScript.StdOut.WriteLine "Done executing tests!"

WScript.Quit 0
