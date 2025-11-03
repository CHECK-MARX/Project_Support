Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing        | Out-Null

function _BA_StyleButton([System.Windows.Forms.Button]$b){
  try{
    # 既存の色味は変えず、枠だけ青系へ
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize  = 1
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80,120,255)
    # ダークテーマでも背景が変わらないように
    $b.FlatAppearance.MouseOverBackColor = $b.BackColor
    $b.FlatAppearance.MouseDownBackColor = $b.BackColor

    if(-not $b.PSObject.Properties['AccentState']){
      $b | Add-Member -NotePropertyName AccentState -NotePropertyValue ([pscustomobject]@{Hover=$false;Focus=$false}) -Force
    }

    $b.add_GotFocus( { param($s,$e) $s.AccentState.Focus=$true;  $s.Invalidate() })
    $b.add_LostFocus({ param($s,$e) $s.AccentState.Focus=$false; $s.Invalidate() })
    $b.add_MouseEnter({ param($s,$e) $s.AccentState.Hover=$true; $s.Invalidate() })
    $b.add_MouseLeave({ param($s,$e) $s.AccentState.Hover=$false;$s.Invalidate() })

    # 光って見える枠を描画（フォーカス/ホバー時のみ）
    $b.add_Paint({
      param($s,$pe)
      $st = $s.AccentState
      if(-not $st){ return }
      if($st.Hover -or $st.Focus){
        $g=$pe.Graphics
        $g.SmoothingMode=[System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $r=[System.Drawing.Rectangle]::Inflate($s.ClientRectangle,-2,-2)

        $outer = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(80, 64,130,255), 4)   # うっすら
        $inner = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(200,64,130,255), 2)   # 明るめ

        try{
          if($st.Hover){ $g.DrawRectangle($outer,$r) }
          $g.DrawRectangle($inner,$r)
        } finally { $outer.Dispose(); $inner.Dispose() }
      }
    }.GetNewClosure())
  }catch{}
}

function Wire-ButtonAccent {
  # フォーム描画完了後に一度だけ全ボタンへ適用
  $idle = $null
  $idle = [System.EventHandler]{
    try{
      $form = [System.Windows.Forms.Application]::OpenForms | Select-Object -First 1
      if(-not $form){ return }
      [System.Windows.Forms.Application]::remove_Idle($idle)

      $stack = New-Object System.Collections.Stack
      $stack.Push($form) | Out-Null
      while($stack.Count){
        $c=$stack.Pop()
        if($c -is [System.Windows.Forms.Button]){ _BA_StyleButton $c }
        foreach($ch in $c.Controls){ $stack.Push($ch) }
      }
    }catch{}
  }
  [System.Windows.Forms.Application]::add_Idle($idle)
}