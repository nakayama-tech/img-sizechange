Option Explicit

Const wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Const cmd = " | .\ConvertTo-Jpeg.ps1"
Dim imgObj							'画像オブジェクト
Dim imgFolderName					'画像ファイル格納場所名
Dim imgFolderObj					'画像ファイル格納場所
Dim imgExt							'画像ファイル拡張子
Dim imgKakuchoushiChangeAfter		'変更後画像ファイル名
Dim imgKakuchoushiChangeAfterObj	'変更後画像ファイルオブジェクト
Dim imgKakuchoushiChangeAfterFolder '変換後画像ファイル格納先
Dim objFS							'ファイルシステムオブジェクト
Dim imgObjOpen						'イメージ（現在）
Dim baseSize						'基本サイズ
Dim objImgChange					'イメージ（変換後）
Dim imgName							'ファイル名
Dim runFolder

Set runFolder = CreateObject("Scripting.FileSystemObject")

imgFolderName = InputBox("画像ファイル格納場所を教えて")

'===============================
' ファイルシステムオブジェクト作成
'===============================
Set objFS = CreateObject("Scripting.FileSystemObject")

'===============================
' フォルダ存在確認
'===============================
If objFS.FolderExists(imgFolderName) Then
	'===============================
	'入力チェック
	'===============================
	baseSize = CInt(InputBox("変換後サイズを入力。例えば1200pxにしたい場合は1200と入力。\r\n変換は縦横大きい方を基準にして変換します"))
	If baseSize <= 0 Then
		'入力値なし、または0の場合
		Set objFS = Nothing
		WScript.Echo "なんかいれて"
		WScript.Quit
	End If
	'===============================
	' 存在する場合はフォルダオブジェクト取得
	'===============================
	Set imgFolderObj = objFS.GetFolder(imgFolderName)
	imgKakuchoushiChangeAfterFolder = imgFolderName & "\henkango"
	If objFS.FolderExists(imgKakuchoushiChangeAfterFolder) Then
	Else
		objFS.createFolder(imgKakuchoushiChangeAfterFolder)
	End If

	'===============================
	' ConvertTo-Jpeg.ps1を使う
	'===============================
	Dim objWshShell
	Set objWshShell = Wscript.CreateObject("WScript.Shell")
	For Each imgObj In imgFolderObj.Files
		imgExt = objFS.GetExtensionName(imgObj.name)
  		If imgExt = "heic" or imgExt = "HEIC" Then
  			'WScript.echo runFolder.getParentFolderName(WScript.ScriptFullName) & "\ConvertTo-Jpeg.ps1 " & imgObj
		  	objWshShell.Run "Powershell -ExecutionPolicy Bypass -NoExit " & runFolder.getParentFolderName(WScript.ScriptFullName) & "\ConvertTo-Jpeg.ps1 " & imgObj, 0,false
		  	WScript.Sleep 1000	' heicファイルが重いことを考慮して1秒待つ
		End If
	Next
	Set objWshShell = Nothing

	WScript.Sleep 3000	' heicファイル終了まで待つ、余裕持って3秒
	'===============================
	' サイズ変更
	'===============================	
	For Each imgObj In imgFolderObj.Files
    	'===============================
		' 拡張子チェック
		'===============================
		imgExt = objFS.GetExtensionName(imgObj.name)
		imgName = imgObj.name
  		If imgExt = "jpg" or imgExt = "JPG" or imgExt = "jpeg" or imgExt = "JPEG" or imgExt = "png" or imgExt = "PNG" or imgExt = "gif" or imgExt = "TIFF" or imgExt = "tiff" or imgExt = "bmp" Then
			'===============================
			' イメージファイルオープン
			'===============================
			Set imgObjOpen = CreateObject("WIA.ImageFile")
			imgObjOpen.LoadFile(imgObj)
			'===============================
			' 縦横1200px以上の場合だけ処理
			'===============================
			If imgObjOpen.Width >= baseSize or imgObjOpen.Height >= baseSize Then
				' 変換後のイメージ
				'WScript.Echo imgName
				Set objImgChange = CreateObject("WIA.ImageProcess")
				'===============================
				' 横か縦、大きい方を1200pxに指定
				'===============================
				If imgObjOpen.Width >= imgObjOpen.Height Then
					'==============================================
					' 画像の大きさ指定（1つ目の要素）※横長
					'==============================================
					objImgChange.Filters.Add(objImgChange.FilterInfos("Scale").FilterID)	' フィルターIDセット
					objImgChange.Filters(1).Properties("MaximumWidth").Value = 1200			' 幅
					objImgChange.Filters(1).Properties("MaximumHeight").Value = imgObjOpen.Height * (1200/imgObjOpen.Width)	' 高さ
				Else
					'==============================================
					' 画像の大きさ指定（1つ目の要素）※長
					'==============================================
					objImgChange.Filters.Add(objImgChange.FilterInfos("Scale").FilterID)	' フィルターIDセット
					objImgChange.Filters(1).Properties("MaximumWidth").Value = imgObjOpen.Width * (1200/imgObjOpen.Height)			' 幅
					objImgChange.Filters(1).Properties("MaximumHeight").Value = 1200	' 高さ
				End If
				'==============================================
				' コンバート情報
				'==============================================
				objImgChange.Filters.Add(objImgChange.FilterInfos("Convert").FilterID)	' フィルターIDセット
				objImgChange.Filters(2).Properties("FormatID").Value = wiaFormatJPEG 
				Set imgObjOpen = objImgChange.Apply(imgObjOpen)	'
				imgObjOpen.SaveFile(objFS.GetAbsolutePathName(imgKakuchoushiChangeAfterFolder & "\" & objFS.getBaseName(imgName) & ".jpg"))

 				Set objImgChange = Nothing

			End If
			
			Set imgObjOpen = Nothing
		
		End If
    Next
    
    set imgFolderObj = Nothing
	set objWshShell = Nothing

Else
	'フォルダが存在しない場合の処理
	WScript.Echo "フォルダがねーぞ"
End If

set objFS = Nothing
set runFolder = Nothing
