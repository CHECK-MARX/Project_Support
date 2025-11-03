# ===== openonly.patch (external) =====
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function _Open_ClearClickHandlers([System.Windows.Forms.Control]$Control){
  try{
    $flags=[System.Reflection.BindingFlags]'Instance,NonPublic,Static'
    $t=[System.Windows.Forms.Control]
    $events=$t.GetProperty('Events',$flags).GetValue($Control,@())
    $key=$t.GetField('EventClick',$flags).GetValue($null)
    $remove=$events.GetType().GetMethod('RemoveHandler',$flags)
    $handlers=$events.Item($key)
    if($handlers){ foreach($d in $handlers.GetInvocationList()){ $remove.Invoke($events,@($key,$d)) } }
  }catch{}
}

function _Open_GetTargetPath($ed){
  try{
    if (Get-Command Get-NoteEditorTargetPath -ErrorAction SilentlyContinue) {
      return Get-NoteEditorTargetPath -Case $script:currentCase -Editor $ed
    }
  }catch{}
  if ($ed -and $ed.FileBox -and $ed.FileBox.Tag) { return [string]$ed.FileBox.Tag }
  return $null
}

function _Open_UpdateButtons(){
  try{
    if (Get-Command Update-NoteEditorButtons -ErrorAction SilentlyContinue) {
      Update-NoteEditorButtons -HasCase:([bool]$script:currentCase); return
    }
  }catch{}
  if(-not $script:noteEditors){ return }
  $has = ($script:currentCase -ne $null)
  foreach($ed in $script:noteEditors){
    if($ed.OpenButton){  $ed.OpenButton.Enabled  = $has }
    if($ed.CloseButton){ $ed.CloseButton.Enabled = $false } # Closeは使わない
  }
}

function Wire-NoteOpenOnly {
  if(-not $script:noteEditors){ return }
  foreach($ed in $script:noteEditors){
    if($ed.OpenButton){
      _Open_ClearClickHandlers $ed.OpenButton
      if(-not ($ed.PSObject.Properties['OpenProcess'])){ $ed | Add-Member -NotePropertyName OpenProcess -NotePropertyValue $null -Force }
      $ed.OpenButton.Tag = $ed
      $ed.OpenButton.add_Click({
        param($sender,$e)
        $x=$sender.Tag; if(-not $x){ return }

        # 案件コンテキスト（既存があれば使う）
        if (Get-Command Ensure-CurrentCaseContext -ErrorAction SilentlyContinue) {
          if(-not (Ensure-CurrentCaseContext -Silent)){ 
            [System.Windows.Forms.MessageBox]::Show('履歴から案件を選択してください。','警告','OK','Warning')|Out-Null
            return
          }
        }

        $path = _Open_GetTargetPath $x
        if(-not $path){
          [System.Windows.Forms.MessageBox]::Show('ファイルを決定できませんでした。','エラー','OK','Error')|Out-Null
          return
        }
        try{
          if(-not (Test-Path -LiteralPath $path)){ Set-Content -LiteralPath $path -Value '' -Encoding UTF8 }
        }catch{
          [System.Windows.Forms.MessageBox]::Show(("ファイルの作成に失敗しました。詳細: {0}" -f $_.Exception.Message),'エラー','OK','Error')|Out-Null
          return
        }

        # 既にプロセス保持がある  まだ生きてればそのまま
        if($x.PSObject.Properties['OpenProcess'] -and $x.OpenProcess){
          try{ if(-not $x.OpenProcess.HasExited){
            [System.Windows.Forms.MessageBox]::Show('このファイルは既に開いています。','情報','OK','Information')|Out-Null
            _Open_UpdateButtons; return
          } }catch{ $x.OpenProcess=$null }
        }

        try{
          $p = Start-Process -FilePath 'notepad.exe' -ArgumentList ("`"{0}`"" -f $path) -PassThru -ErrorAction Stop
          $x.OpenProcess = $p
          if($x.FileBox){ $x.FileBox.Text = [IO.Path]::GetFileName($path); $x.FileBox.Tag = $path }
        }catch{
          [System.Windows.Forms.MessageBox]::Show(("ファイルを開けませんでした。詳細: {0}" -f $_.Exception.Message),'エラー','OK','Error')|Out-Null
          return
        }
        _Open_UpdateButtons
      }.GetNewClosure())
    }

    # Closeボタンは無効化（誤作動防止）
    if($ed.CloseButton){
      _Open_ClearClickHandlers $ed.CloseButton
      $ed.CloseButton.Enabled = $false
    }
  }
  _Open_UpdateButtons
}

# 互換: 既存コードが Wire-NoteOpenClose を呼んでもOKにする
function Wire-NoteOpenClose { Wire-NoteOpenOnly }
# ===== end openonly.patch =====