Dim App 'As Application
Set App = CreateObject("QuickTest.Application")
'App.Launch
'App.Visible = True
App.Options.DisableVORecognition = False
App.Options.AutoGenerateWith = False
App.Options.WithGenerationLevel = 2
App.Options.TimeToActivateWinAfterPoint = 500
App.Options.SaveLoadAndMonitorData = False
App.Options.Run.RunMode = "Normal"
App.Options.Run.ViewResults = False
App.Options.Run.CaptureForTestResults = "Never"
App.Options.Run.StepExecutionDelay = 0
App.Options.TE.CurrentEmulator = "IBM PCom 5.7"
App.Options.TE.Protocol = "autodetect"
App.Options.TE.AutoAdvance = 0
App.Options.TE.CodePage = 0
App.Options.TE.HllapiDllName = "C:\Program Files\IBM\Personal Communications\pcshll32.dll"
App.Options.TE.HllapiProcName = "hllapi"
App.Options.TE.VerifyHllapiDllPath = 1
App.Options.TE.AutoSyncKeys = "13"
App.Options.TE.RecordMenusAndPopups = 1
App.Options.TE.RecordCursorPosition = 1
App.Options.TE.TrailingMode = 1
App.Options.TE.TrailingFieldLength = 5
App.Options.TE.UsePropertyPattern = 1
App.Options.TE.PropertyPatternsFile = "C:\Program Files\Mercury Interactive\QuickTest Professional\dat\PropertyPatternConfigTE.xml"
App.Options.TE.SyncTime = 200
App.Options.TE.ScreenTitleRow = "1"
App.Options.TE.ScreenTitleCol = "1"
App.Options.TE.ScreenTitleLength = "30"
App.Options.WindowsApps.AttachedTextRadius = 35
App.Options.WindowsApps.AttachedTextArea = "TopLeft"
App.Options.WindowsApps.ExpandMenuToRetrieveProperties = True
App.Options.WindowsApps.NonUniqueListItemRecordMode = "ByName"
App.Options.WindowsApps.RecordOwnerDrawnButtonAs = "PushButtons"
App.Options.WindowsApps.ForceEnumChildWindows = 0
App.Options.WindowsApps.ClickEditBeforeSetText = 0
App.Options.WindowsApps.VerifyMenuInitEvent = 0
App.Folders.RemoveAll
