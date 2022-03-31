' =========================変数=============================
strFName = !ファイル名!
strMacroName = !マクロ名!
blnVBA = !VBA用|True,False!
vn_result = $マクロ実行結果$
' =========================変数=============================

If strFName = "" Then
  Err.Raise 1, "", "指定されたファイルを開くことができません。" 
End If

If strMacroName = "" Then
  Err.Raise 1, "", "VBAマクロ名を設定してください。" 
End If

Set objAcsApp = Wscript.GetObject(strFName)

If blnVBA Then
 result=objAcsApp.Run(strMacroName)
Else
 result=objAcsApp.DoCmd.RunMacro(strMacroName)
End If

' return
SetUmsVariable vn_result,result

Set objAcsApp = nothing
