Add-Type -AssemblyName System.Windows.Forms | Out-Null

function _DM_Path() {
  $cfg = Join-Path $PSScriptRoot '..\config'
  try{ if(-not (Test-Path $cfg)){ New-Item -ItemType Directory -Path $cfg -Force | Out-Null } }catch{}
  Join-Path $cfg 'user-settings.json'
}
function _DM_Load() {
  $p=_DM_Path
  try{
    if(Test-Path $p){
      $t=Get-Content -LiteralPath $p -Raw -Encoding UTF8
      if(-not [string]::IsNullOrWhiteSpace($t)){ return $t | ConvertFrom-Json }
    }
  }catch{}
  [pscustomobject]@{ DarkMode=$null }
}
function _DM_Save([bool]$on){
  try{ ([pscustomobject]@{ DarkMode=[bool]$on } | ConvertTo-Json) | Set-Content -LiteralPath (_DM_Path) -Encoding UTF8 }catch{}
}
function _DM_FindDarkCheck([System.Windows.Forms.Form]$form){
  if(-not $form){ return $null }
  $nameHints = '^chk.*dark|^cb.*dark|darkmode|theme.*dark'
  $textHints = '^\s*ダーク\s*モード\s*$|^\s*ダーク\s*テーマ\s*$|Dark\s*Mode|Dark\s*Theme'
  $stack = New-Object System.Collections.Stack; $stack.Push($form) | Out-Null
  while($stack.Count){
    $c=$stack.Pop()
    if($c -is [System.Windows.Forms.CheckBox]){
      if( ($c.Name -match $nameHints) -or ($c.Text -match $textHints) ){ return $c }
    }
    foreach($ch in $c.Controls){ $stack.Push($ch) }
  }
  $null
}
function _DM_ApplyTheme([bool]$on){
  foreach($name in 'Apply-DarkMode','Set-DarkMode','UpdateTheme','Set-Theme'){
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if($cmd){ try{ & $cmd -Dark:$on -On:$on -DarkMode:$on -Enable:$on -ErrorAction SilentlyContinue }catch{}; break }
  }
}

function Wire-PersistDarkMode-Min() {
  $idle = $null
  $idle = [System.EventHandler]{
    try{
      $form = [System.Windows.Forms.Application]::OpenForms | Select-Object -First 1
      if(-not $form){ return }
      [System.Windows.Forms.Application]::remove_Idle($idle)

      $chk = _DM_FindDarkCheck $form
      if(-not $chk){ return }

      # 起動時：保存値がある時だけ一度だけ反映（ユーザー操作を上書きしない）
      $cfg = _DM_Load
      if($cfg.DarkMode -ne $null){
        $applying=$true
        try{ $chk.Checked = [bool]$cfg.DarkMode } finally { $applying=$false }
        _DM_ApplyTheme ([bool]$cfg.DarkMode)
      }

      # 変更時：即保存＆反映
      $chk.add_CheckedChanged({
        if($applying){ return }
        _DM_Save ([bool]$chk.Checked)
        _DM_ApplyTheme ([bool]$chk.Checked)
      }.GetNewClosure())

      # 終了時：念のため最終状態も保存
      $form.add_FormClosing({ try{ _DM_Save ([bool]$chk.Checked) }catch{} }.GetNewClosure())
    }catch{}
  }
  [System.Windows.Forms.Application]::add_Idle($idle)
}