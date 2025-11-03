# ===== persist-darkmode.patch (robust) =====
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function _Dm_SettingsPath {
  try {
    $cfgDir = Join-Path $PSScriptRoot '..\config'
    if(-not (Test-Path $cfgDir)){ New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null }
    return (Join-Path $cfgDir 'user-settings.json')
  } catch { return "$PSScriptRoot\..\config\user-settings.json" }
}

function _Dm_Load(){
  $p=_Dm_SettingsPath
  try{
    if(Test-Path -LiteralPath $p){
      $txt = Get-Content -LiteralPath $p -Raw -Encoding UTF8
      if(-not [string]::IsNullOrWhiteSpace($txt)){ return $txt | ConvertFrom-Json }
    }
  }catch{}
  [pscustomobject]@{ DarkMode = $false }
}

function _Dm_Save([bool]$val){
  try{ ([pscustomobject]@{ DarkMode=[bool]$val } | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath (_Dm_SettingsPath) -Encoding UTF8 }catch{}
}

function _Dm_ApplyIfAny([bool]$on){
  foreach($name in 'Set-DarkMode','Apply-DarkMode','UpdateTheme','Set-Theme'){
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if($cmd){ try{ & $cmd -Dark:$on -On:$on -Enable:$on -DarkMode:$on -ErrorAction SilentlyContinue }catch{}; break }
  }
}

function _Dm_FindCheckBox([System.Windows.Forms.Control]$root){
  if(-not $root){ return $null }
  $names = 'chkDark','chkDarkMode','DarkMode','darkModeCheck','chkThemeDark'
  $stack = New-Object System.Collections.Stack; $stack.Push($root)|Out-Null
  $best=$null
  while($stack.Count -gt 0){
    $c=$stack.Pop()
    if($c -is [System.Windows.Forms.CheckBox]){
      if($names -contains $c.Name){ return $c }
      $t=[string]$c.Text
      if($t -match 'ダーク|Dark|テーマ.*(暗|dark)'){ return $c }
      if(-not $best){ $best=$c }
    }
    foreach($ch in $c.Controls){ $stack.Push($ch) }
  }
  return $best
}

function _Dm_EnumToolStripItems([System.Windows.Forms.ToolStripItemCollection]$items){
  foreach($it in $items){
    $it
    if($it -is [System.Windows.Forms.ToolStripDropDownItem]){
      foreach($sub in (_Dm_EnumToolStripItems $it.DropDownItems)){ $sub }
    }
  }
}

function _Dm_FindToolStripToggle([System.Windows.Forms.Form]$form){
  $candidates = @()
  if($form.MainMenuStrip){
    $candidates += _Dm_EnumToolStripItems $form.MainMenuStrip.Items
  }
  foreach($ctl in $form.Controls){
    if($ctl -is [System.Windows.Forms.MenuStrip]){
      $candidates += _Dm_EnumToolStripItems $ctl.Items
    }
    if($ctl -is [System.Windows.Forms.ToolStrip]){
      $candidates += _Dm_EnumToolStripItems $ctl.Items
    }
  }
  $names='miDark','tsmiDark','menuDark','darkModeMenu','menuThemeDark'
  foreach($it in $candidates){
    if($names -contains $it.Name){ return $it }
    $t=[string]$it.Text
    if($t -match 'ダーク|Dark|テーマ.*(暗|dark)'){ return $it }
  }
  return $null
}

function Wire-PersistDarkMode {
  $form = [System.Windows.Forms.Application]::OpenForms | Select-Object -First 1
  if(-not $form){ return }

  # 1) UI要素を探す（CheckBox 優先、なければ ToolStrip のチェックメニュー）
  $chk = _Dm_FindCheckBox $form
  $ts  = _Dm_FindToolStripToggle $form

  if(-not $chk -and -not $ts){ return }

  $cfg = _Dm_Load
  $state = [bool]$cfg.DarkMode

  $apply = {
    param($on)
    if($args.Count -gt 0){ $on = [bool]$args[0] } else { $on = $state }
    if($chk){ try{ $script:_dm_applying=$true; $chk.Checked = $on } finally { $script:_dm_applying=$false } }
    if($ts){  try{ $ts.Checked = $on }catch{} }
    _Dm_ApplyIfAny $on
  }

  # 2) Shown 後に適用（起動時の既存コードによる上書きを回避）
  $once = $false
  $form.add_Shown({
    if($once){ return } $once=$true
    & $apply $state
    # 念押し：100ms 後にもう一度（遅延初期化に勝つ）
    $t = New-Object System.Windows.Forms.Timer
    $t.Interval = 120
    $t.add_Tick({ $t.Stop(); try{ & $apply $state }catch{}; try{$t.Dispose()}catch{} })
    $t.Start()
  }.GetNewClosure())

  # 3) 変更時に保存
  if($chk){
    $chk.add_CheckedChanged({
      if($script:_dm_applying){ return }
      $state = [bool]$this.Checked
      _Dm_Save $state
      _Dm_ApplyIfAny $state
    }.GetNewClosure())
  }
  if($ts){
    $ts.add_Click({
      $state = -not [bool]$ts.Checked  # クリック直後はまだトグル前の値のことがあるので反転想定
      try{ $ts.Checked = $state }catch{}
      _Dm_Save $state
      _Dm_ApplyIfAny $state
    }.GetNewClosure())
  }

  # 4) 終了時に最終保存
  $form.add_FormClosing({ try{ _Dm_Save $state }catch{} }.GetNewClosure())
}
# ===== end persist-darkmode.patch =====