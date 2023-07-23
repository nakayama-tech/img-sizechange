Option Explicit

Const wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Const cmd = " | .\ConvertTo-Jpeg.ps1"
Dim imgObj							'�摜�I�u�W�F�N�g
Dim imgFolderName					'�摜�t�@�C���i�[�ꏊ��
Dim imgFolderObj					'�摜�t�@�C���i�[�ꏊ
Dim imgExt							'�摜�t�@�C���g���q
Dim imgKakuchoushiChangeAfter		'�ύX��摜�t�@�C����
Dim imgKakuchoushiChangeAfterObj	'�ύX��摜�t�@�C���I�u�W�F�N�g
Dim imgKakuchoushiChangeAfterFolder '�ϊ���摜�t�@�C���i�[��
Dim objFS							'�t�@�C���V�X�e���I�u�W�F�N�g
Dim imgObjOpen						'�C���[�W�i���݁j
Dim baseSize						'��{�T�C�Y
Dim objImgChange					'�C���[�W�i�ϊ���j
Dim imgName							'�t�@�C����
Dim runFolder

Set runFolder = CreateObject("Scripting.FileSystemObject")

imgFolderName = InputBox("�摜�t�@�C���i�[�ꏊ��������")

'===============================
' �t�@�C���V�X�e���I�u�W�F�N�g�쐬
'===============================
Set objFS = CreateObject("Scripting.FileSystemObject")

'===============================
' �t�H���_���݊m�F
'===============================
If objFS.FolderExists(imgFolderName) Then
	'===============================
	'���̓`�F�b�N
	'===============================
	baseSize = CInt(InputBox("�ϊ���T�C�Y����́B�Ⴆ��1200px�ɂ������ꍇ��1200�Ɠ��́B\r\n�ϊ��͏c���傫��������ɂ��ĕϊ����܂�"))
	If baseSize <= 0 Then
		'���͒l�Ȃ��A�܂���0�̏ꍇ
		Set objFS = Nothing
		WScript.Echo "�Ȃ񂩂����"
		WScript.Quit
	End If
	'===============================
	' ���݂���ꍇ�̓t�H���_�I�u�W�F�N�g�擾
	'===============================
	Set imgFolderObj = objFS.GetFolder(imgFolderName)
	imgKakuchoushiChangeAfterFolder = imgFolderName & "\henkango"
	If objFS.FolderExists(imgKakuchoushiChangeAfterFolder) Then
	Else
		objFS.createFolder(imgKakuchoushiChangeAfterFolder)
	End If

	'===============================
	' ConvertTo-Jpeg.ps1���g��
	'===============================
	Dim objWshShell
	Set objWshShell = Wscript.CreateObject("WScript.Shell")
	For Each imgObj In imgFolderObj.Files
		imgExt = objFS.GetExtensionName(imgObj.name)
  		If imgExt = "heic" or imgExt = "HEIC" Then
  			'WScript.echo runFolder.getParentFolderName(WScript.ScriptFullName) & "\ConvertTo-Jpeg.ps1 " & imgObj
		  	objWshShell.Run "Powershell -ExecutionPolicy Bypass -NoExit " & runFolder.getParentFolderName(WScript.ScriptFullName) & "\ConvertTo-Jpeg.ps1 " & imgObj, 0,false
		  	WScript.Sleep 1000	' heic�t�@�C�����d�����Ƃ��l������1�b�҂�
		End If
	Next
	Set objWshShell = Nothing

	WScript.Sleep 3000	' heic�t�@�C���I���܂ő҂A�]�T������3�b
	'===============================
	' �T�C�Y�ύX
	'===============================	
	For Each imgObj In imgFolderObj.Files
    	'===============================
		' �g���q�`�F�b�N
		'===============================
		imgExt = objFS.GetExtensionName(imgObj.name)
		imgName = imgObj.name
  		If imgExt = "jpg" or imgExt = "JPG" or imgExt = "jpeg" or imgExt = "JPEG" or imgExt = "png" or imgExt = "PNG" or imgExt = "gif" or imgExt = "TIFF" or imgExt = "tiff" or imgExt = "bmp" Then
			'===============================
			' �C���[�W�t�@�C���I�[�v��
			'===============================
			Set imgObjOpen = CreateObject("WIA.ImageFile")
			imgObjOpen.LoadFile(imgObj)
			'===============================
			' �c��1200px�ȏ�̏ꍇ��������
			'===============================
			If imgObjOpen.Width >= baseSize or imgObjOpen.Height >= baseSize Then
				' �ϊ���̃C���[�W
				'WScript.Echo imgName
				Set objImgChange = CreateObject("WIA.ImageProcess")
				'===============================
				' �����c�A�傫������1200px�Ɏw��
				'===============================
				If imgObjOpen.Width >= imgObjOpen.Height Then
					'==============================================
					' �摜�̑傫���w��i1�ڂ̗v�f�j������
					'==============================================
					objImgChange.Filters.Add(objImgChange.FilterInfos("Scale").FilterID)	' �t�B���^�[ID�Z�b�g
					objImgChange.Filters(1).Properties("MaximumWidth").Value = 1200			' ��
					objImgChange.Filters(1).Properties("MaximumHeight").Value = imgObjOpen.Height * (1200/imgObjOpen.Width)	' ����
				Else
					'==============================================
					' �摜�̑傫���w��i1�ڂ̗v�f�j����
					'==============================================
					objImgChange.Filters.Add(objImgChange.FilterInfos("Scale").FilterID)	' �t�B���^�[ID�Z�b�g
					objImgChange.Filters(1).Properties("MaximumWidth").Value = imgObjOpen.Width * (1200/imgObjOpen.Height)			' ��
					objImgChange.Filters(1).Properties("MaximumHeight").Value = 1200	' ����
				End If
				'==============================================
				' �R���o�[�g���
				'==============================================
				objImgChange.Filters.Add(objImgChange.FilterInfos("Convert").FilterID)	' �t�B���^�[ID�Z�b�g
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
	'�t�H���_�����݂��Ȃ��ꍇ�̏���
	WScript.Echo "�t�H���_���ˁ[��"
End If

set objFS = Nothing
set runFolder = Nothing
