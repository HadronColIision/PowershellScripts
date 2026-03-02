param([switch]$MacroMode)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Threading;
public static class Win32 {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern bool FreeConsole();
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

if ($MacroMode) {
    [Win32]::FreeConsole() | Out-Null

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
        [Win32]::keybd_event($script:hcKey1, 0, 0, [UIntPtr]::Zero)
        [System.Threading.Thread]::Sleep($script:delays.hcPress_swap)
        [Win32]::keybd_event($script:hcKey1, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        [System.Threading.Thread]::Sleep($script:delays.hcPress_afterSwap)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.hcPress_click)
        if ($script:hcAction2 -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::keybd_event($script:hcKey2, 0, 0, [UIntPtr]::Zero)
            [System.Threading.Thread]::Sleep($script:delays.hcPress_key2)
            [Win32]::keybd_event($script:hcKey2, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        }
        [System.Threading.Thread]::Sleep($script:delays.hcPress_end)
    }

    function Run-HitCrystal-WhileHolding {
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.hcHold_click)
        [Win32]::LeftClick()
    }

    function Run-SingleAnchor {
        [Win32]::PressKey(0x43)
        [System.Threading.Thread]::Sleep($script:delays.sa_crystal)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.sa_place)
        [Win32]::PressKey($script:ancSlot)
        [System.Threading.Thread]::Sleep($script:delays.sa_slot)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.sa_use)
        if ($script:ancBackAction -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::PressKey($script:ancBackKey)
        }
        [System.Threading.Thread]::Sleep($script:delays.sa_back)
        [Win32]::RightClick()
    }

    function Run-DoubleAnchor {
        [Win32]::PressKey(0x43)
        [System.Threading.Thread]::Sleep($script:delays.da_crystal1)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place1)
        [Win32]::PressKey($script:ancSlot)
        [System.Threading.Thread]::Sleep($script:delays.da_slot1)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_use1)
        [Win32]::PressKey(0x43)
        [System.Threading.Thread]::Sleep($script:delays.da_crystal2)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place2)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place3)
        [Win32]::PressKey($script:ancSlot)
        [System.Threading.Thread]::Sleep($script:delays.da_slot2)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_use2)
        if ($script:ancBackAction -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::PressKey($script:ancBackKey)
        }
        [System.Threading.Thread]::Sleep($script:delays.da_back)
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

    $script:activePreset  = "Brxtwurst"
    $script:hcKey1        = 0x32
    $script:hcAction2     = "key"
    $script:hcKey2        = 0x33
    $script:ancSlot       = 0x56
    $script:ancBackAction = "mouse"
    $script:ancBackKey    = 0x00
    $script:presetBrxBtn  = $null
    $script:presetWanBtn  = $null
    $script:delays = @{
        hcPress_swap      = 2
        hcPress_afterSwap = 1
        hcPress_click     = 25
        hcPress_key2      = 10
        hcPress_end       = 1
        hcHold_click      = 35
        sa_crystal        = 10
        sa_place          = 25
        sa_slot           = 20
        sa_use            = 10
        sa_back           = 10
        da_crystal1       = 15
        da_place1         = 25
        da_slot1          = 25
        da_use1           = 25
        da_crystal2       = 15
        da_place2         = 1
        da_place3         = 25
        da_slot2          = 25
        da_use2           = 25
        da_back           = 15
    }

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
                $script:keyBtns[$ci].BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
            }
        } catch {}
    }

    function New-BrxtwurstForm {
        $w = 360
        $h = 340
        $pad = 20
        $contentW = $w - ($pad * 2)

        # --- Colors (Catppuccin Mocha) ---
        $baseBg     = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
        $mantleBg   = [System.Drawing.ColorTranslator]::FromHtml("#181825")
        $surface0   = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        $surface1   = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
        $overlay0   = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
        $textClr    = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
        $accentClr  = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
        $greenClr   = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
        $redClr     = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
        $yellowClr  = [System.Drawing.ColorTranslator]::FromHtml("#f9e2af")

        $f = New-Object System.Windows.Forms.Form
        $f.Text            = "Wanda Macros"
        $f.ClientSize      = New-Object System.Drawing.Size($w, $h)
        $f.StartPosition   = "CenterScreen"
        $f.FormBorderStyle = "None"
        $f.BackColor       = $baseBg
        $f.Font            = New-Object System.Drawing.Font("Segoe UI", 10)
        $f.TopMost         = $true
        $f.KeyPreview      = $true
        $f.ShowInTaskbar   = $true
        $f.MinimizeBox     = $true
        $f.Region          = [System.Drawing.Region]::new([System.Drawing.Rectangle]::new(0, 0, $w, $h))

        $f.Add_KeyDown({
            if ($script:listening -ge 0) {
                Finish-Listen $_.KeyCode.ToString()
                $_.Handled = $true
                $_.SuppressKeyPress = $true
            }
        })

        # Drag support
        $f.Add_MouseDown({ if ($_.Button -eq "Left") { $this.Tag = $_.Location } })
        $f.Add_MouseMove({
            if ($_.Button -eq "Left" -and $this.Tag) {
                $p = $this.Tag
                $this.Location = New-Object System.Drawing.Point(
                    ($this.Location.X + $_.X - $p.X),
                    ($this.Location.Y + $_.Y - $p.Y))
            }
        })
        $f.Add_MouseUp({ $this.Tag = $null })

        # Accent bar
        $accent = New-Object System.Windows.Forms.Panel
        $accent.Size      = New-Object System.Drawing.Size($w, 3)
        $accent.Location  = New-Object System.Drawing.Point(0, 0)
        $accent.BackColor = $accentClr
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

        # Title
        $titleLbl = New-Object System.Windows.Forms.Label
        $titleLbl.Text      = "Wanda Macros"
        $titleLbl.ForeColor = $textClr
        $titleLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 13)
        $titleLbl.AutoSize  = $false
        $titleLbl.Size      = New-Object System.Drawing.Size(200, 36)
        $titleLbl.Location  = New-Object System.Drawing.Point($pad, 8)
        $titleLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $titleLbl.Add_MouseDown({ if ($_.Button -eq "Left") { $this.FindForm().Tag = $_.Location } })
        $titleLbl.Add_MouseMove({
            if ($_.Button -eq "Left" -and $this.FindForm().Tag) {
                $p = $this.FindForm().Tag
                $this.FindForm().Location = New-Object System.Drawing.Point(
                    ($this.FindForm().Location.X + $_.X - $p.X),
                    ($this.FindForm().Location.Y + $_.Y - $p.Y))
            }
        })
        $titleLbl.Add_MouseUp({ $this.FindForm().Tag = $null })
        $f.Controls.Add($titleLbl)

        # Version
        $verLbl = New-Object System.Windows.Forms.Label
        $verLbl.Text      = "v0.4"
        $verLbl.ForeColor = $surface0
        $verLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
        $verLbl.AutoSize  = $true
        $verLbl.Location  = New-Object System.Drawing.Point(165, 18)
        $f.Controls.Add($verLbl)

        # Minimize
        $minBtn = New-Object System.Windows.Forms.Label
        $minBtn.Text      = [char]0x2015
        $minBtn.ForeColor = $overlay0
        $minBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 11)
        $minBtn.AutoSize  = $true
        $minBtn.Location  = New-Object System.Drawing.Point(($w - 52), 8)
        $minBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $minBtn.Add_Click({ $this.FindForm().WindowState = [System.Windows.Forms.FormWindowState]::Minimized })
        $minBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7") })
        $minBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $f.Controls.Add($minBtn)

        # Close
        $xBtn = New-Object System.Windows.Forms.Label
        $xBtn.Text      = [char]0x2715
        $xBtn.ForeColor = $overlay0
        $xBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 11)
        $xBtn.AutoSize  = $true
        $xBtn.Location  = New-Object System.Drawing.Point(($w - 28), 8)
        $xBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $xBtn.Add_Click({ $this.FindForm().Hide() })
        $xBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8") })
        $xBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $f.Controls.Add($xBtn)

        $cy = 50

        # Separator
        $sep = New-Object System.Windows.Forms.Panel
        $sep.Size      = New-Object System.Drawing.Size($contentW, 1)
        $sep.Location  = New-Object System.Drawing.Point($pad, $cy)
        $sep.BackColor = $surface0
        $f.Controls.Add($sep)
        $cy += 14

        # Preset buttons
        $presetW = [int](($contentW - 8) / 2)
        $presetH = 32

        $brxBtn = New-Object System.Windows.Forms.Button
        $brxBtn.Text      = "Brxtwurst"
        $brxBtn.Size      = New-Object System.Drawing.Size($presetW, $presetH)
        $brxBtn.Location  = New-Object System.Drawing.Point($pad, $cy)
        $brxBtn.FlatStyle = "Flat"
        $brxBtn.FlatAppearance.BorderSize         = 1
        $brxBtn.FlatAppearance.MouseOverBackColor = $surface0
        $brxBtn.FlatAppearance.MouseDownBackColor = $accentClr
        $brxBtn.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $brxBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        if ($script:activePreset -eq "Brxtwurst") {
            $brxBtn.BackColor = $surface0; $brxBtn.ForeColor = $accentClr
            $brxBtn.FlatAppearance.BorderColor = $accentClr
        } else {
            $brxBtn.BackColor = $mantleBg; $brxBtn.ForeColor = $surface1
            $brxBtn.FlatAppearance.BorderColor = $surface0
        }
        $brxBtn.Add_Click({
            $script:activePreset  = "Brxtwurst"
            $script:hcKey1        = 0x32; $script:hcAction2 = "key"; $script:hcKey2 = 0x33
            $script:ancSlot       = 0x56; $script:ancBackAction = "mouse"; $script:ancBackKey = 0x00
            $ac = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
            $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
            $this.ForeColor = $ac; $this.FlatAppearance.BorderColor = $ac
            $script:presetWanBtn.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
            $script:presetWanBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
            $script:presetWanBtn.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        })
        $f.Controls.Add($brxBtn)
        $script:presetBrxBtn = $brxBtn

        $wanBtn = New-Object System.Windows.Forms.Button
        $wanBtn.Text      = "Wanda"
        $wanBtn.Size      = New-Object System.Drawing.Size($presetW, $presetH)
        $wanBtn.Location  = New-Object System.Drawing.Point(($pad + $presetW + 8), $cy)
        $wanBtn.FlatStyle = "Flat"
        $wanBtn.FlatAppearance.BorderSize         = 1
        $wanBtn.FlatAppearance.MouseOverBackColor = $surface0
        $wanBtn.FlatAppearance.MouseDownBackColor = $accentClr
        $wanBtn.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $wanBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        if ($script:activePreset -eq "Wanda") {
            $wanBtn.BackColor = $surface0; $wanBtn.ForeColor = $accentClr
            $wanBtn.FlatAppearance.BorderColor = $accentClr
        } else {
            $wanBtn.BackColor = $mantleBg; $wanBtn.ForeColor = $surface1
            $wanBtn.FlatAppearance.BorderColor = $surface0
        }
        $wanBtn.Add_Click({
            $script:activePreset  = "Wanda"
            $script:hcKey1        = 0x51; $script:hcAction2 = "mouse"; $script:hcKey2 = 0x00
            $script:ancSlot       = 0x31; $script:ancBackAction = "key"; $script:ancBackKey = 0x33
            $ac = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
            $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
            $this.ForeColor = $ac; $this.FlatAppearance.BorderColor = $ac
            $script:presetBrxBtn.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
            $script:presetBrxBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
            $script:presetBrxBtn.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        })
        $f.Controls.Add($wanBtn)
        $script:presetWanBtn = $wanBtn
        $cy += $presetH + 16

        # Macro rows
        $macroNames = @("Hit Crystal", "Single Anchor", "Double Anchor")
        $kbW = 72
        $toggleW = 56
        $rowH = 40

        for ($idx = 0; $idx -lt 3; $idx++) {
            # Row background
            $row = New-Object System.Windows.Forms.Panel
            $row.Size      = New-Object System.Drawing.Size($contentW, $rowH)
            $row.Location  = New-Object System.Drawing.Point($pad, $cy)
            $row.BackColor = $mantleBg
            $f.Controls.Add($row)

            # Status dot
            $dot = New-Object System.Windows.Forms.Panel
            $dot.Size     = New-Object System.Drawing.Size(8, 8)
            $dot.Location = New-Object System.Drawing.Point(12, 16)
            if ($script:btnStates[$idx]) { $dot.BackColor = $greenClr } else { $dot.BackColor = $surface1 }
            $row.Controls.Add($dot)
            $script:dotRefs[$idx] = $dot

            # Name
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text      = $macroNames[$idx]
            $lbl.ForeColor = $textClr
            $lbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
            $lbl.AutoSize  = $false
            $lbl.Size      = New-Object System.Drawing.Size(($contentW - $toggleW - $kbW - 50), $rowH)
            $lbl.Location  = New-Object System.Drawing.Point(28, 0)
            $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $row.Controls.Add($lbl)

            # Toggle
            $b = New-Object System.Windows.Forms.Button
            if ($script:btnStates[$idx]) { $b.Text = "ON" } else { $b.Text = "OFF" }
            $b.Size      = New-Object System.Drawing.Size($toggleW, 28)
            $b.Location  = New-Object System.Drawing.Point(($contentW - $toggleW - $kbW - 12), 6)
            $b.FlatStyle = "Flat"
            $b.FlatAppearance.BorderSize         = 1
            $b.FlatAppearance.MouseOverBackColor = $surface0
            $b.FlatAppearance.MouseDownBackColor = $accentClr
            $b.BackColor = $baseBg
            $b.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
            $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $b.Tag       = $idx
            if ($script:btnStates[$idx]) {
                $b.ForeColor = $greenClr; $b.FlatAppearance.BorderColor = $greenClr
            } else {
                $b.ForeColor = $surface1; $b.FlatAppearance.BorderColor = $surface0
            }
            $b.Add_Click({
                $ci = $this.Tag
                $script:btnStates[$ci] = -not $script:btnStates[$ci]
                $gn = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
                $gy = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
                $bd = [System.Drawing.ColorTranslator]::FromHtml("#313244")
                if ($script:btnStates[$ci]) {
                    $this.Text = "ON"; $this.ForeColor = $gn; $this.FlatAppearance.BorderColor = $gn
                    $script:dotRefs[$ci].BackColor = $gn
                } else {
                    $this.Text = "OFF"; $this.ForeColor = $gy; $this.FlatAppearance.BorderColor = $bd
                    $script:dotRefs[$ci].BackColor = $gy
                }
            })
            $row.Controls.Add($b)
            $script:btnRefs[$idx] = $b

            # Keybind
            $kb = New-Object System.Windows.Forms.Button
            $kb.Size      = New-Object System.Drawing.Size($kbW, 28)
            $kb.Location  = New-Object System.Drawing.Point(($contentW - $kbW - 6), 6)
            $kb.FlatStyle = "Flat"
            $kb.FlatAppearance.BorderSize         = 1
            $kb.FlatAppearance.BorderColor        = $surface0
            $kb.FlatAppearance.MouseOverBackColor = $surface0
            $kb.FlatAppearance.MouseDownBackColor = $accentClr
            $kb.BackColor = $baseBg
            $kb.ForeColor = $surface1
            $kb.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
            $kb.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $kb.Tag       = $idx
            if ($script:btnKeys[$idx] -eq "None") { $kb.Text = "..." }
            else { $kb.Text = $script:btnKeys[$idx]; $kb.ForeColor = $textClr }
            $kb.Add_Click({
                $ci = $this.Tag
                if ($script:listening -ge 0 -and $script:listening -ne $ci) {
                    try {
                        $prev = $script:listening
                        if ($script:btnKeys[$prev] -eq "None") {
                            $script:keyBtns[$prev].Text = "..."
                            $script:keyBtns[$prev].ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
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
            $row.Controls.Add($kb)
            $script:keyBtns[$idx] = $kb

            $dot.BringToFront()
            $cy += $rowH + 6
        }

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
    $tray.Text    = "Wanda Macros - Home to toggle"
    $tray.Visible = $true

    $bmp = New-Object System.Drawing.Bitmap(16,16)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.Clear([System.Drawing.ColorTranslator]::FromHtml("#cba6f7"))
    $gfx.Dispose()
    $tray.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $exit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit Wanda Macros")
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

    $script:guiForm = New-BrxtwurstForm

    $appCtx = New-Object System.Windows.Forms.ApplicationContext
    [System.Windows.Forms.Application]::Run($appCtx)
} else {
    $ErrorActionPreference = 'Stop'

    try {
        # Kill any existing macro instances
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

        # Launch macros in background
        $scriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $null }
        if ($scriptPath -and (Test-Path $scriptPath -ErrorAction SilentlyContinue)) {
            Start-Process -WindowStyle Hidden -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
                -ArgumentList "-ExecutionPolicy Bypass -STA -NoProfile -File `"$scriptPath`" -MacroMode"
        }

        # Run HabibiModAnalyzer in foreground
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        Invoke-Expression (Invoke-RestMethod "https://raw.githubusercontent.com/HadronCollision/PowershellScripts/refs/heads/main/HabibiModAnalyzer.ps1")
    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
    }
}
