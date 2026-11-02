<# themes.psm1 #>

$rawThemes = @{
    OriginalGray = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(30, 30, 30)     # Form backgrounds ($mainForm, $settingsForm, tabGeneral)
        SecondaryBg                 = @(40, 40, 40)     # Input panels, DataGrids, control surfaces
        TopBarBg                    = @(20, 20, 20)     # TopBar and bottom ResizeGrip panel background
        BorderColor                 = @(100, 100, 100)  # ToolTip outline and default dividers
        PanelBg                     = @(35, 35, 38)     # MasterHotkeysPanel container

        # --- Typography & Content ---
        PrimaryText                 = @(240, 240, 240)  # Standard label, button, and control text
        HeaderTitleText             = @(240, 240, 240)  # TitleLabel foreground
        HeaderSubtitleText          = @(230, 230, 230)  # CopyrightLabel foreground
        SectionHeaderText           = @(160, 160, 160)  # Category subheaders (e.g., BASIC SETTINGS)
        AccentDanger                = @(255, 46, 76)    # ExtraForm AccentColor / Report alert labels
        StatusActive                = @(100, 200, 100)  # MasterHotkeysStatus & ChestStatusLabel
        CurrencyPerin               = @(97, 175, 239)   # UpgradeTotalPerin label
        CurrencyEcp                 = @(152, 195, 121)  # UpgradeTotalEcp label

        # --- Semantic Aliases ---
        WhiteText           = @(240, 240, 240)
        SecondaryText       = @(230, 230, 230)
        AccentBlue          = @(35, 90, 180)
        TertiaryBg          = @(35, 35, 38)
        TopBarAccent        = @(35, 90, 180)
        InputBgAlt          = @(40, 40, 40)
        SuccessColor        = @(100, 200, 100)
        SuccessColorAlt     = @(100, 200, 100)
        WarningColor        = @(255, 180, 0)
        WarningColorAlt     = @(255, 180, 0)
        ErrorColor          = @(255, 46, 76)
        ErrorColorAlt       = @(255, 46, 76)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(139, 0, 0)
        StatusDisconnectedText    = @(0, 0, 0)
        StatusFlashing            = @(255, 165, 0)
        StatusFlashingText        = @(0, 0, 0)
        StatusFlashingSelect      = @(255, 140, 0)
        NotificationProgressHigh  = @(0, 180, 0)
        NotificationProgressMid   = @(220, 180, 0)
        NotificationProgressLow   = @(200, 50, 50)
        NotificationStripWarning  = @(255, 165, 0)
        NotificationStripInfo     = @(100, 149, 237)
        NotificationStripError    = @(205, 92, 92)
        NotificationBg            = @(50, 50, 50)
        NotificationStripBg       = @(70, 70, 70)
        NotificationBtnBg         = @(80, 80, 80)
        NotificationBtnHover      = @(100, 100, 100)
        NotificationBtnPressed    = @(60, 60, 60)
        NotificationMsgText       = @(211, 211, 211)
        TooltipBorderColor        = @(100, 100, 100)

        # --- ToolTips ---
        ToolTipBg                   = @(30, 30, 30)     # ToolTip SolidBrush canvas
        ToolTipBorder               = @(100, 100, 100)  # ToolTip border Pen
        ToolTipFg                   = @(240, 240, 240)  # ToolTip text brush

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(40, 40, 40)     # Standard button background ($btnLaunch, $btnLogin, etc.)
        HeaderButtonBg              = @(40, 40, 40)     # Window control buttons ($btnInfoForm, $btnMinimizeForm)
        BossButtonBg                = @(70, 36, 36)     # World Boss action buttons ($btnNormal)
        BossTimerText               = @(255, 255, 255)  # Boss timer countdown overlay text

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(40, 40, 40)     # Standard TextBox & ComboBox background
        InputFg                     = @(240, 240, 240)  # Standard TextBox & ComboBox text
        ComboDropdownBg             = @(40, 40, 40)     # Dropdown list canvas
        ComboItemText               = @(240, 240, 240)  # Dropdown item text
        ComboBorderNormal           = @(100, 100, 100)  # Unfocused outline
        ComboArrow                  = @(240, 240, 240)  # Dropdown arrow glyph
        ComboSelectedItem           = @(35, 90, 180)    # Active selection highlight
        ComboSelectedText           = @(255, 255, 255)  # Active selection text

        # --- DataGridView & Tables ---
        GridBg                      = @(40, 40, 40)     # DataGridMain background
        GridBorder                  = @(60, 60, 60)     # Grid cell border lines
        GridSecondaryBg             = @(40, 40, 40)     # Default row background
        GridAlternateRow            = @(35, 35, 35)     # Alternating row background
        GridSelectedRow             = @(35, 90, 180)    # Selected row background
        GridSelectedText            = @(255, 255, 255)  # Selected row text
        GridHeaderBg                = @(30, 30, 30)     # Column header surface
        GridHeaderText              = @(240, 240, 240)  # Column header text
        GridRowHeaderBg             = @(30, 30, 30)     # Row selector column background
        GridRowHeaderText           = @(240, 240, 240)  # Row selector text

        # --- Tab Controls & Pages ---
        TabBg                       = @(30, 30, 30)     # DarkTabControl header strip background
        TabActive                   = @(40, 40, 40)     # Active tab header background
        TabInactive                 = @(30, 30, 30)     # Inactive tab header background
        TabActiveText               = @(240, 240, 240)  # Active tab text
        TabInactiveText             = @(160, 160, 160)  # Inactive tab text
        TabUnderline                = @(35, 90, 180)    # Active tab bottom accent underline
        TabInactiveUnderline        = @(50, 50, 50)     # Tab strip bottom divider
        TabPageBg                   = @(30, 30, 30)     # TabPage surface ($tabGeneral, $tabLoginSettings, etc.)
        TabPageText                 = @(240, 240, 240)  # TabPage content text

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(35, 175, 75)    # Green button/switch indicator
        ToggleOffColor              = @(210, 45, 45)    # Red cancel/switch indicator
        GroupBoxBg                  = @(35, 35, 38)     # Panel/GroupBox container surface
        GroupBoxText                = @(240, 240, 240)  # GroupBox title
        ListItemBg                  = @(40, 40, 40)     # List item canvas
        ListItemText                = @(240, 240, 240)  # List item text
        ProgressBarBg               = @(40, 40, 40)     # Progress bar track
        ProgressBarFill             = @(35, 90, 180)    # Progress bar indicator
        TreeViewBg                  = @(22, 26, 38)     # Secondary dark surface (ExtraForm/Notes/Queues)
        TreeViewHover               = @(50, 50, 50)     # Node hover highlight
        ScrollThumbNormal           = @(50, 50, 50)     # ScrollBar slider thumb
        ScrollThumbHover            = @(70, 70, 70)     # ScrollBar slider hover
        ScrollThumbActive           = @(90, 90, 90)     # ScrollBar slider dragging

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(50, 50, 50)     # $grayNormalStart
        BtnGradientNormalBot        = @(35, 35, 35)     # $grayNormalEnd
        BtnGradientHoverTop         = @(70, 70, 70)     # $grayHoverStart
        BtnGradientHoverBot         = @(50, 50, 50)     # $grayHoverEnd
        BtnGradientLeftPressTop     = @(90, 90, 90)     # $grayLeftStart
        BtnGradientLeftPressBot     = @(70, 70, 70)     # $grayLeftEnd
        BtnGradientRightPressTop    = @(80, 80, 110)    # $grayRightStart
        BtnGradientRightPressBot    = @(55, 55, 80)     # $grayRightEnd

        BtnDangerGradientNormalTop    = @(170, 30, 30)  # $redNormalStart
        BtnDangerGradientNormalBot    = @(130, 15, 15)  # $redNormalEnd
        BtnDangerGradientHoverTop     = @(200, 40, 40)  # $redHoverStart
        BtnDangerGradientHoverBot     = @(155, 20, 20)  # $redHoverEnd
        BtnDangerGradientLeftPressTop = @(140, 15, 15)  # $redLeftStart
        BtnDangerGradientLeftPressBot = @(100, 10, 10)  # $redLeftEnd
        BtnDangerGradientRightPressTop= @(180, 30, 90)  # $redRightStart
        BtnDangerGradientRightPressBot= @(145, 15, 65)  # $redRightEnd

        BtnSuccessGradientNormalTop    = @(40, 140, 70) # $greenNormalStart
        BtnSuccessGradientNormalBot    = @(25, 95, 45)  # $greenNormalEnd
        BtnSuccessGradientHoverTop     = @(55, 175, 90) # $greenHoverStart
        BtnSuccessGradientHoverBot     = @(40, 140, 70) # $greenHoverEnd
        BtnSuccessGradientLeftPressTop = @(30, 110, 55) # $greenLeftStart
        BtnSuccessGradientLeftPressBot = @(20, 75, 35)  # $greenLeftEnd
        BtnSuccessGradientRightPressTop= @(60, 150, 85) # $greenRightStart
        BtnSuccessGradientRightPressBot= @(45, 110, 60) # $greenRightEnd

        BtnPrimaryGradientNormalTop    = @(35, 90, 180) # $blueNormalStart
        BtnPrimaryGradientNormalBot    = @(20, 60, 130) # $blueNormalEnd
        BtnPrimaryGradientHoverTop     = @(50, 115, 220)# $blueHoverStart
        BtnPrimaryGradientHoverBot     = @(35, 90, 180) # $blueHoverEnd
        BtnPrimaryGradientLeftPressTop = @(25, 70, 140) # $blueLeftStart
        BtnPrimaryGradientLeftPressBot = @(15, 45, 100) # $blueLeftEnd
        BtnPrimaryGradientRightPressTop= @(55, 100, 190)# $blueRightStart
        BtnPrimaryGradientRightPressBot= @(40, 75, 145) # $blueRightEnd
    }
	
	Dark = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(13, 17, 23)     # Form backgrounds, CheckBox surfaces, ScrollBar track
        SecondaryBg                 = @(22, 27, 34)     # Panels and container surfaces
        TopBarBg                    = @(13, 17, 23)     # Navigation top bar gradient bottom stop
        BorderColor                 = @(48, 54, 61)     # Dividers, borders & control outlines
        PanelBg                     = @(22, 27, 34)     # Generic container panels

        # --- Typography & Content ---
        PrimaryText                 = @(230, 237, 243)  # Main text for forms, buttons, checkboxes, grids
        HeaderTitleText             = @(230, 237, 243)  # Title headings
        HeaderSubtitleText          = @(139, 148, 158)  # Subtitles, helper text & muted labels
        SectionHeaderText           = @(88, 166, 255)   # Sub-headers and section category titles
        AccentDanger                = @(255, 46, 76)    # Red highlights, warnings, and report accents
        StatusActive                = @(63, 185, 80)    # Active status indicators
        CurrencyPerin               = @(97, 175, 239)   # Perin totals
        CurrencyEcp                 = @(152, 195, 121)  # ECP totals

        # --- Semantic Aliases ---
        WhiteText           = @(230, 237, 243)
        SecondaryText       = @(139, 148, 158)
        AccentBlue          = @(31, 111, 235)
        TertiaryBg          = @(33, 38, 45)
        TopBarAccent        = @(88, 166, 255)
        InputBgAlt          = @(22, 27, 34)
        SuccessColor        = @(63, 185, 80)
        SuccessColorAlt     = @(63, 185, 80)
        WarningColor        = @(255, 180, 0)
        WarningColorAlt     = @(255, 180, 0)
        ErrorColor          = @(255, 46, 76)
        ErrorColorAlt       = @(255, 46, 76)

        # --- ToolTips ---
        ToolTipBg                   = @(22, 27, 34)     # ToolTip canvas
        ToolTipBorder               = @(48, 54, 61)     # ToolTip outline
        ToolTipFg                   = @(230, 237, 243)  # ToolTip text

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(33, 38, 45)     # Default button surface
        HeaderButtonBg              = @(33, 38, 45)     # Top bar gradient top stop
        BossButtonBg                = @(70, 36, 36)     # World boss action buttons
        BossTimerText               = @(255, 255, 255)  # World boss countdown text overlay

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(22, 27, 34)     # Standard TextBox and ComboBox background
        InputFg                     = @(230, 237, 243)  # Standard TextBox and ComboBox foreground
        ComboDropdownBg             = @(22, 27, 34)     # Custom.DarkComboBox dropdown surface
        ComboItemText               = @(230, 237, 243)  # Custom.DarkComboBox item text
        ComboBorderNormal           = @(48, 54, 61)     # Custom.DarkComboBox unfocused border

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(139, 0, 0)
        StatusDisconnectedText    = @(0, 0, 0)
        StatusFlashing            = @(255, 165, 0)
        StatusFlashingText        = @(0, 0, 0)
        StatusFlashingSelect      = @(255, 140, 0)
        NotificationProgressHigh  = @(40, 180, 80)
        NotificationProgressMid   = @(240, 190, 40)
        NotificationProgressLow   = @(220, 60, 60)
        NotificationStripWarning  = @(255, 180, 0)
        NotificationStripInfo     = @(70, 130, 220)
        NotificationStripError    = @(220, 70, 70)
        NotificationBg            = @(31, 34, 42)
        NotificationStripBg       = @(48, 54, 61)
        NotificationBtnBg         = @(48, 54, 61)
        NotificationBtnHover      = @(66, 72, 82)
        NotificationBtnPressed    = @(22, 26, 32)
        NotificationMsgText       = @(189, 197, 207)
        TooltipBorderColor        = @(76, 84, 96)
        ComboArrow                  = @(230, 237, 243)  # Custom.DarkComboBox chevron glyph
        ComboSelectedItem           = @(31, 111, 235)   # Custom.DarkComboBox selected item background
        ComboSelectedText           = @(255, 255, 255)  # Custom.DarkComboBox selected item text

        # --- DataGridView & Tables ---
        GridBg                      = @(13, 17, 23)     # DataGridView overall background canvas
        GridBorder                  = @(48, 54, 61)     # Grid line borders
        GridSecondaryBg             = @(22, 27, 34)     # DataGridView default cell row background
        GridAlternateRow            = @(28, 33, 40)     # Alternating zebra-stripe row background
        GridSelectedRow             = @(31, 111, 235)   # Selected row / cell highlight
        GridSelectedText            = @(255, 255, 255)  # Selected row / cell text
        GridHeaderBg                = @(33, 38, 45)     # Column header surface
        GridHeaderText              = @(240, 240, 240)  # Column header title text
        GridRowHeaderBg             = @(33, 38, 45)     # Row selector background
        GridRowHeaderText           = @(240, 240, 240)  # Row selector text

        # --- Tab Controls & Pages ---
        TabBg                       = @(13, 17, 23)     # DarkTabControl & TabControl background
        TabActive                   = @(22, 27, 34)     # Active tab header background
        TabInactive                 = @(13, 17, 23)     # Inactive tab header background
        TabActiveText               = @(230, 237, 243)  # Active tab header text
        TabInactiveText             = @(139, 148, 158)  # Inactive tab header text
        TabUnderline                = @(31, 111, 235)   # Active tab underline indicator
        TabInactiveUnderline        = @(48, 54, 61)     # Inactive tab bottom baseline
        TabPageBg                   = @(13, 17, 23)     # TabPage body surface
        TabPageText                 = @(230, 237, 243)  # TabPage body text

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(35, 175, 75)    # Custom.Toggle active ON switch color
        ToggleOffColor              = @(210, 45, 45)    # Custom.Toggle inactive OFF switch color
        GroupBoxBg                  = @(22, 27, 34)     # GroupBox background
        GroupBoxText                = @(230, 237, 243)  # GroupBox caption text
        ListItemBg                  = @(22, 27, 34)     # ListBox background
        ListItemText                = @(230, 237, 243)  # ListBox item text
        ProgressBarBg               = @(22, 27, 34)     # ProgressBar empty track background
        ProgressBarFill             = @(31, 111, 235)   # ProgressBar filled bar color
        TreeViewBg                  = @(22, 26, 38)     # TreeView background surface
        TreeViewHover               = @(45, 50, 65)     # TreeView node hover highlight
        ScrollThumbNormal           = @(33, 38, 45)     # Custom.DarkScrollBar thumb normal state
        ScrollThumbHover            = @(48, 54, 61)     # Custom.DarkScrollBar thumb hover state
        ScrollThumbActive           = @(88, 166, 255)   # Custom.DarkScrollBar thumb dragging state

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(42, 48, 58)
        BtnGradientNormalBot        = @(28, 33, 40)
        BtnGradientHoverTop         = @(62, 70, 84)
        BtnGradientHoverBot         = @(48, 55, 67)
        BtnGradientLeftPressTop     = @(30, 36, 46)
        BtnGradientLeftPressBot     = @(20, 24, 32)
        BtnGradientRightPressTop    = @(68, 58, 110)
        BtnGradientRightPressBot    = @(50, 42, 85)

        BtnDangerGradientNormalTop    = @(185, 42, 42)
        BtnDangerGradientNormalBot    = @(140, 22, 22)
        BtnDangerGradientHoverTop     = @(215, 55, 55)
        BtnDangerGradientHoverBot     = @(165, 30, 30)
        BtnDangerGradientLeftPressTop = @(140, 22, 22)
        BtnDangerGradientLeftPressBot = @(100, 14, 14)
        BtnDangerGradientRightPressTop= @(185, 42, 90)
        BtnDangerGradientRightPressBot= @(140, 22, 65)

        BtnSuccessGradientNormalTop    = @(48, 155, 75)
        BtnSuccessGradientNormalBot    = @(30, 108, 50)
        BtnSuccessGradientHoverTop     = @(60, 180, 92)
        BtnSuccessGradientHoverBot     = @(48, 155, 75)
        BtnSuccessGradientLeftPressTop = @(34, 120, 58)
        BtnSuccessGradientLeftPressBot = @(22, 80, 38)
        BtnSuccessGradientRightPressTop= @(58, 160, 88)
        BtnSuccessGradientRightPressBot= @(44, 118, 64)

        BtnPrimaryGradientNormalTop    = @(48, 100, 195)
        BtnPrimaryGradientNormalBot    = @(30, 68, 145)
        BtnPrimaryGradientHoverTop     = @(64, 125, 230)
        BtnPrimaryGradientHoverBot     = @(48, 100, 195)
        BtnPrimaryGradientLeftPressTop = @(30, 70, 150)
        BtnPrimaryGradientLeftPressBot = @(18, 46, 108)
        BtnPrimaryGradientRightPressTop= @(58, 100, 200)
        BtnPrimaryGradientRightPressBot= @(42, 76, 155)
    }

    Light = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(240, 240, 240)
        SecondaryBg                 = @(245, 245, 245)
        TopBarBg                    = @(240, 240, 240)
        BorderColor                 = @(180, 180, 180)
        PanelBg                     = @(245, 245, 245)

        # --- Typography & Content ---
        PrimaryText                 = @(40, 40, 40)
        HeaderTitleText             = @(20, 20, 20)
        HeaderSubtitleText          = @(110, 110, 110)
        SectionHeaderText           = @(9, 105, 218)
        AccentDanger                = @(207, 34, 46)
        StatusActive                = @(26, 127, 55)
        CurrencyPerin               = @(13, 102, 208)
        CurrencyEcp                 = @(40, 140, 50)

        # --- Semantic Aliases ---
        WhiteText           = @(255, 255, 255)
        SecondaryText       = @(110, 110, 110)
        AccentBlue          = @(9, 105, 218)
        TertiaryBg          = @(235, 235, 235)
        TopBarAccent        = @(0, 120, 215)
        InputBgAlt          = @(250, 250, 250)
        SuccessColor        = @(26, 127, 55)
        SuccessColorAlt     = @(26, 127, 55)
        WarningColor        = @(180, 120, 0)
        WarningColorAlt     = @(180, 120, 0)
        ErrorColor          = @(207, 34, 46)
        ErrorColorAlt       = @(207, 34, 46)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(180, 180, 180)
        ToolTipFg                   = @(40, 40, 40)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(225, 225, 225)
        HeaderButtonBg              = @(225, 225, 225)
        BossButtonBg                = @(245, 215, 215)
        BossTimerText               = @(40, 40, 40)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(40, 40, 40)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(40, 40, 40)
        ComboBorderNormal           = @(180, 180, 180)
        ComboArrow                  = @(90, 90, 90)
        ComboSelectedItem           = @(0, 120, 215)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(250, 250, 250)
        GridBorder                  = @(200, 200, 200)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(245, 245, 245)
        GridSelectedRow             = @(0, 120, 215)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(225, 225, 225)
        GridHeaderText              = @(40, 40, 40)
        GridRowHeaderBg             = @(225, 225, 225)
        GridRowHeaderText           = @(40, 40, 40)

        # --- Tab Controls & Pages ---
        TabBg                       = @(240, 240, 240)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(180, 30, 30)
        StatusDisconnectedText    = @(255, 255, 255)
        StatusFlashing            = @(255, 140, 0)
        StatusFlashingText        = @(0, 0, 0)
        StatusFlashingSelect      = @(210, 105, 0)
        NotificationProgressHigh  = @(34, 160, 50)
        NotificationProgressMid   = @(220, 170, 30)
        NotificationProgressLow   = @(200, 60, 60)
        NotificationStripWarning  = @(255, 140, 0)
        NotificationStripInfo     = @(70, 110, 200)
        NotificationStripError    = @(180, 70, 70)
        NotificationBg            = @(240, 240, 240)
        NotificationStripBg       = @(200, 200, 200)
        NotificationBtnBg         = @(200, 200, 200)
        NotificationBtnHover      = @(180, 180, 180)
        NotificationBtnPressed    = @(170, 170, 170)
        NotificationMsgText       = @(80, 80, 80)
        TooltipBorderColor        = @(180, 180, 180)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(240, 240, 240)
        TabActiveText               = @(40, 40, 40)
        TabInactiveText             = @(120, 120, 120)
        TabUnderline                = @(0, 120, 215)
        TabInactiveUnderline        = @(200, 200, 200)
        TabPageBg                   = @(240, 240, 240)
        TabPageText                 = @(40, 40, 40)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(0, 180, 80)
        ToggleOffColor              = @(180, 80, 80)
        GroupBoxBg                  = @(245, 245, 245)
        GroupBoxText                = @(40, 40, 40)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(40, 40, 40)
        ProgressBarBg               = @(230, 230, 230)
        ProgressBarFill             = @(0, 120, 215)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(220, 220, 230)
        ScrollThumbNormal           = @(180, 180, 180)
        ScrollThumbHover            = @(160, 160, 160)
        ScrollThumbActive           = @(0, 120, 215)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(245, 245, 245)
        BtnGradientNormalBot        = @(225, 225, 225)
        BtnGradientHoverTop         = @(235, 235, 235)
        BtnGradientHoverBot         = @(215, 215, 215)
        BtnGradientLeftPressTop     = @(200, 200, 200)
        BtnGradientLeftPressBot     = @(190, 190, 190)
        BtnGradientRightPressTop    = @(210, 200, 230)
        BtnGradientRightPressBot    = @(190, 180, 210)

        BtnDangerGradientNormalTop    = @(230, 120, 120)
        BtnDangerGradientNormalBot    = @(220, 100, 100)
        BtnDangerGradientHoverTop     = @(240, 130, 130)
        BtnDangerGradientHoverBot     = @(220, 100, 100)
        BtnDangerGradientLeftPressTop = @(210, 90, 90)
        BtnDangerGradientLeftPressBot = @(190, 70, 70)
        BtnDangerGradientRightPressTop= @(220, 100, 140)
        BtnDangerGradientRightPressBot= @(200, 80, 120)

        BtnSuccessGradientNormalTop    = @(120, 210, 120)
        BtnSuccessGradientNormalBot    = @(100, 200, 100)
        BtnSuccessGradientHoverTop     = @(140, 220, 140)
        BtnSuccessGradientHoverBot     = @(110, 210, 110)
        BtnSuccessGradientLeftPressTop = @(90, 190, 90)
        BtnSuccessGradientLeftPressBot = @(70, 170, 70)
        BtnSuccessGradientRightPressTop= @(110, 200, 120)
        BtnSuccessGradientRightPressBot= @(90, 180, 100)

        BtnPrimaryGradientNormalTop    = @(100, 170, 240)
        BtnPrimaryGradientNormalBot    = @(0, 120, 215)
        BtnPrimaryGradientHoverTop     = @(120, 190, 250)
        BtnPrimaryGradientHoverBot     = @(50, 150, 230)
        BtnPrimaryGradientLeftPressTop = @(0, 100, 190)
        BtnPrimaryGradientLeftPressBot = @(0, 80, 170)
        BtnPrimaryGradientRightPressTop= @(40, 110, 210)
        BtnPrimaryGradientRightPressBot= @(20, 90, 190)
    }

    # =========================================================================
    # 1. DEEP OCEAN (Dark Mode)
    # =========================================================================
    DeepOcean = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(8, 15, 26)       # Deepest ocean trench
        SecondaryBg                 = @(14, 25, 42)      # Elevated panel surface
        TopBarBg                    = @(6, 12, 20)       # Recessed top navigation
        BorderColor                 = @(24, 44, 72)      # Marine navy border
        PanelBg                     = @(14, 25, 42)

        # --- Typography & Content ---
        PrimaryText                 = @(230, 242, 255)   # Crisp seafoam white
        HeaderTitleText             = @(245, 250, 255)
        HeaderSubtitleText          = @(125, 155, 190)   # Muted slate mist
        SectionHeaderText           = @(56, 189, 248)    # Bright cyan
        AccentDanger                = @(251, 113, 133)   # Soft coral rose
        StatusActive                = @(52, 211, 153)    # Marine emerald
        CurrencyPerin               = @(56, 189, 248)    # Cyan crystal
        CurrencyEcp                 = @(129, 140, 248)   # Periwinkle

        # --- Semantic Aliases ---
        WhiteText           = @(230, 242, 255)
        SecondaryText       = @(125, 155, 190)
        AccentBlue          = @(56, 189, 248)
        TertiaryBg          = @(20, 36, 60)
        TopBarAccent        = @(56, 189, 248)
        InputBgAlt          = @(14, 25, 42)
        SuccessColor        = @(52, 211, 153)
        SuccessColorAlt     = @(52, 211, 153)
        WarningColor        = @(250, 200, 60)
        WarningColorAlt     = @(250, 200, 60)
        ErrorColor          = @(251, 113, 133)
        ErrorColorAlt       = @(251, 113, 133)

        # --- ToolTips ---
        ToolTipBg                   = @(11, 20, 34)
        ToolTipBorder               = @(56, 189, 248)
        ToolTipFg                   = @(230, 242, 255)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(20, 36, 60)
        HeaderButtonBg              = @(14, 25, 42)
        BossButtonBg                = @(56, 24, 38)
        BossTimerText               = @(255, 235, 240)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(10, 18, 30)      # Recessed input well
        InputFg                     = @(230, 242, 255)
        ComboDropdownBg             = @(14, 25, 42)
        ComboItemText               = @(230, 242, 255)
        ComboBorderNormal           = @(30, 56, 92)
        ComboArrow                  = @(56, 189, 248)
        ComboSelectedItem           = @(14, 116, 204)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(8, 15, 26)
        GridBorder                  = @(20, 38, 64)
        GridSecondaryBg             = @(14, 25, 42)
        GridAlternateRow            = @(11, 20, 34)
        GridSelectedRow             = @(14, 116, 204)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(20, 36, 60)
        GridHeaderText              = @(230, 242, 255)
        GridRowHeaderBg             = @(20, 36, 60)
        GridRowHeaderText           = @(230, 242, 255)

        # --- Tab Controls & Pages ---
        TabBg                       = @(8, 15, 26)
        TabActive                   = @(14, 25, 42)
        TabInactive                 = @(8, 15, 26)
        TabActiveText               = @(230, 242, 255)
        TabInactiveText             = @(125, 155, 190)
        TabUnderline                = @(56, 189, 248)
        TabInactiveUnderline        = @(24, 44, 72)
        TabPageBg                   = @(8, 15, 26)
        TabPageText                 = @(230, 242, 255)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(56, 189, 248)
        ToggleOffColor              = @(70, 30, 42)
        GroupBoxBg                  = @(14, 25, 42)
        GroupBoxText                = @(230, 242, 255)
        ListItemBg                  = @(14, 25, 42)
        ListItemText                = @(230, 242, 255)
        ProgressBarBg               = @(10, 18, 30)
        ProgressBarFill             = @(14, 116, 204)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(139, 20, 20)
        StatusDisconnectedText    = @(220, 220, 240)
        StatusFlashing            = @(255, 180, 40)
        StatusFlashingText        = @(10, 18, 30)
        StatusFlashingSelect      = @(210, 145, 20)
        NotificationProgressHigh  = @(14, 180, 140)
        NotificationProgressMid   = @(255, 180, 40)
        NotificationProgressLow   = @(220, 60, 80)
        NotificationStripWarning  = @(0, 180, 220)
        NotificationStripInfo     = @(56, 140, 220)
        NotificationStripError    = @(200, 60, 80)
        NotificationBg            = @(10, 18, 30)
        NotificationStripBg       = @(24, 44, 72)
        NotificationBtnBg         = @(24, 44, 72)
        NotificationBtnHover      = @(36, 64, 104)
        NotificationBtnPressed    = @(8, 12, 22)
        NotificationMsgText       = @(180, 200, 225)
        TooltipBorderColor        = @(56, 100, 160)
        TreeViewBg                  = @(10, 18, 30)
        TreeViewHover               = @(24, 44, 72)
        ScrollThumbNormal           = @(24, 44, 72)
        ScrollThumbHover            = @(36, 64, 104)
        ScrollThumbActive           = @(56, 189, 248)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(26, 46, 76)
        BtnGradientNormalBot        = @(16, 28, 48)
        BtnGradientHoverTop         = @(36, 64, 104)
        BtnGradientHoverBot         = @(24, 42, 70)
        BtnGradientLeftPressTop     = @(14, 24, 40)
        BtnGradientLeftPressBot     = @(10, 18, 30)
        BtnGradientRightPressTop    = @(30, 52, 86)
        BtnGradientRightPressBot    = @(20, 36, 60)

        BtnDangerGradientNormalTop    = @(190, 45, 75)
        BtnDangerGradientNormalBot    = @(145, 25, 52)
        BtnDangerGradientHoverTop     = @(215, 60, 95)
        BtnDangerGradientHoverBot     = @(170, 32, 65)
        BtnDangerGradientLeftPressTop = @(130, 20, 45)
        BtnDangerGradientLeftPressBot = @(95, 14, 32)
        BtnDangerGradientRightPressTop= @(170, 40, 75)
        BtnDangerGradientRightPressBot= @(130, 25, 58)

        BtnSuccessGradientNormalTop    = @(24, 150, 110)
        BtnSuccessGradientNormalBot    = @(14, 110, 80)
        BtnSuccessGradientHoverTop     = @(36, 178, 130)
        BtnSuccessGradientHoverBot     = @(22, 145, 105)
        BtnSuccessGradientLeftPressTop = @(12, 95, 70)
        BtnSuccessGradientLeftPressBot = @(8, 70, 50)
        BtnSuccessGradientRightPressTop= @(32, 150, 118)
        BtnSuccessGradientRightPressBot= @(18, 115, 90)

        BtnPrimaryGradientNormalTop    = @(20, 135, 225)
        BtnPrimaryGradientNormalBot    = @(12, 95, 170)
        BtnPrimaryGradientHoverTop     = @(45, 165, 250)
        BtnPrimaryGradientHoverBot     = @(18, 130, 220)
        BtnPrimaryGradientLeftPressTop = @(10, 85, 150)
        BtnPrimaryGradientLeftPressBot = @(6, 62, 115)
        BtnPrimaryGradientRightPressTop= @(32, 125, 210)
        BtnPrimaryGradientRightPressBot= @(16, 90, 160)
    }

    # =========================================================================
    # 2. FOREST SAGE (Dark Mode)
    # =========================================================================
    ForestSage = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(12, 20, 15)      # Deep Nordic pine
        SecondaryBg                 = @(20, 34, 26)      # Soft woodland moss panel
        TopBarBg                    = @(8, 15, 11)
        BorderColor                 = @(36, 58, 44)      # Foliage shadow border
        PanelBg                     = @(20, 34, 26)

        # --- Typography & Content ---
        PrimaryText                 = @(235, 245, 238)   # Pale sage cream
        HeaderTitleText             = @(245, 252, 248)
        HeaderSubtitleText          = @(135, 168, 148)   # Muted spruce green
        SectionHeaderText           = @(134, 239, 172)   # Fresh spring bud
        AccentDanger                = @(248, 113, 113)   # Soft rowanberry red
        StatusActive                = @(74, 222, 128)    # Living emerald
        CurrencyPerin               = @(110, 231, 183)   # Mint crystal
        CurrencyEcp                 = @(250, 204, 21)    # Golden amber

        # --- Semantic Aliases ---
        WhiteText           = @(235, 245, 238)
        SecondaryText       = @(130, 160, 140)
        AccentBlue          = @(80, 160, 120)
        TertiaryBg          = @(25, 42, 33)
        TopBarAccent        = @(80, 160, 120)
        InputBgAlt          = @(20, 34, 26)
        SuccessColor        = @(90, 180, 100)
        SuccessColorAlt     = @(90, 180, 100)
        WarningColor        = @(220, 180, 60)
        WarningColorAlt     = @(220, 180, 60)
        ErrorColor          = @(210, 90, 80)
        ErrorColorAlt       = @(210, 90, 80)

        # --- ToolTips ---
        ToolTipBg                   = @(16, 28, 21)
        ToolTipBorder               = @(74, 222, 128)
        ToolTipFg                   = @(235, 245, 238)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(28, 46, 36)
        HeaderButtonBg              = @(20, 34, 26)
        BossButtonBg                = @(62, 28, 30)
        BossTimerText               = @(255, 230, 232)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(14, 24, 18)
        InputFg                     = @(235, 245, 238)
        ComboDropdownBg             = @(20, 34, 26)
        ComboItemText               = @(235, 245, 238)
        ComboBorderNormal           = @(44, 72, 54)
        ComboArrow                  = @(134, 239, 172)
        ComboSelectedItem           = @(34, 130, 80)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(12, 20, 15)
        GridBorder                  = @(28, 48, 36)
        GridSecondaryBg             = @(20, 34, 26)
        GridAlternateRow            = @(16, 28, 21)
        GridSelectedRow             = @(34, 130, 80)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(28, 46, 36)
        GridHeaderText              = @(235, 245, 238)
        GridRowHeaderBg             = @(28, 46, 36)
        GridRowHeaderText           = @(235, 245, 238)

        # --- Tab Controls & Pages ---
        TabBg                       = @(12, 20, 15)
        TabActive                   = @(20, 34, 26)
        TabInactive                 = @(12, 20, 15)
        TabActiveText               = @(235, 245, 238)
        TabInactiveText             = @(135, 168, 148)
        TabUnderline                = @(74, 222, 128)
        TabInactiveUnderline        = @(36, 58, 44)
        TabPageBg                   = @(12, 20, 15)
        TabPageText                 = @(235, 245, 238)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(74, 222, 128)
        ToggleOffColor              = @(65, 30, 32)
        GroupBoxBg                  = @(20, 34, 26)
        GroupBoxText                = @(235, 245, 238)
        ListItemBg                  = @(20, 34, 26)
        ListItemText                = @(235, 245, 238)
        ProgressBarBg               = @(14, 24, 18)
        ProgressBarFill             = @(46, 160, 100)
        TreeViewBg                  = @(14, 24, 18)
        TreeViewHover               = @(32, 54, 42)
        ScrollThumbNormal           = @(36, 58, 44)
        ScrollThumbHover            = @(50, 82, 62)
        ScrollThumbActive           = @(74, 222, 128)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(36, 60, 46)
        BtnGradientNormalBot        = @(24, 40, 30)
        BtnGradientHoverTop         = @(48, 80, 62)
        BtnGradientHoverBot         = @(32, 54, 40)
        BtnGradientLeftPressTop     = @(20, 34, 26)
        BtnGradientLeftPressBot     = @(14, 24, 18)
        BtnGradientRightPressTop    = @(42, 68, 52)
        BtnGradientRightPressBot    = @(28, 46, 36)

        BtnDangerGradientNormalTop    = @(185, 50, 50)
        BtnDangerGradientNormalBot    = @(140, 30, 30)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(155, 40, 30)
        StatusDisconnectedText    = @(230, 235, 225)
        StatusFlashing            = @(220, 175, 50)
        StatusFlashingText        = @(35, 42, 30)
        StatusFlashingSelect      = @(190, 145, 30)
        NotificationProgressHigh  = @(80, 170, 60)
        NotificationProgressMid   = @(200, 170, 40)
        NotificationProgressLow   = @(190, 60, 50)
        NotificationStripWarning  = @(200, 170, 40)
        NotificationStripInfo     = @(80, 150, 120)
        NotificationStripError    = @(180, 70, 60)
        NotificationBg            = @(35, 42, 30)
        NotificationStripBg       = @(50, 62, 42)
        NotificationBtnBg         = @(55, 68, 46)
        NotificationBtnHover      = @(68, 82, 56)
        NotificationBtnPressed    = @(28, 34, 24)
        NotificationMsgText       = @(200, 210, 190)
        TooltipBorderColor        = @(80, 100, 68)
        BtnDangerGradientHoverTop     = @(210, 68, 68)
        BtnDangerGradientHoverBot     = @(165, 40, 40)
        BtnDangerGradientLeftPressTop = @(125, 24, 24)
        BtnDangerGradientLeftPressBot = @(90, 16, 16)
        BtnDangerGradientRightPressTop= @(168, 52, 52)
        BtnDangerGradientRightPressBot= @(128, 32, 32)

        BtnSuccessGradientNormalTop    = @(52, 160, 102)
        BtnSuccessGradientNormalBot    = @(34, 115, 70)
        BtnSuccessGradientHoverTop     = @(68, 190, 124)
        BtnSuccessGradientHoverBot     = @(46, 142, 90)
        BtnSuccessGradientLeftPressTop = @(28, 98, 60)
        BtnSuccessGradientLeftPressBot = @(18, 72, 44)
        BtnSuccessGradientRightPressTop= @(56, 162, 106)
        BtnSuccessGradientRightPressBot= @(38, 120, 78)

        BtnPrimaryGradientNormalTop    = @(46, 148, 95)
        BtnPrimaryGradientNormalBot    = @(28, 105, 66)
        BtnPrimaryGradientHoverTop     = @(62, 178, 116)
        BtnPrimaryGradientHoverBot     = @(38, 130, 82)
        BtnPrimaryGradientLeftPressTop = @(24, 88, 54)
        BtnPrimaryGradientLeftPressBot = @(16, 64, 38)
        BtnPrimaryGradientRightPressTop= @(50, 145, 96)
        BtnPrimaryGradientRightPressBot= @(32, 108, 70)
    }

    # =========================================================================
    # 3. CRIMSON TWILIGHT (Dark Mode)
    # =========================================================================
    CrimsonTwilight = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(20, 13, 16)      # Obsidian cherry
        SecondaryBg                 = @(30, 20, 24)      # Dark plum velvet
        TopBarBg                    = @(14, 9, 11)
        BorderColor                 = @(60, 36, 44)      # Muted wine border
        PanelBg                     = @(30, 20, 24)

        # --- Typography & Content ---
        PrimaryText                 = @(250, 238, 242)   # Rose quartz white
        HeaderTitleText             = @(255, 245, 248)
        HeaderSubtitleText          = @(175, 138, 148)   # Heather mauve
        SectionHeaderText           = @(251, 113, 133)   # Bright petal rose
        AccentDanger                = @(244, 63, 94)     # Vivid crimson
        StatusActive                = @(52, 211, 153)    # Pale jade mint
        CurrencyPerin               = @(251, 113, 133)   # Rose rubies
        CurrencyEcp                 = @(251, 146, 60)    # Sunset peach

        # --- Semantic Aliases ---
        WhiteText           = @(245, 235, 240)
        SecondaryText       = @(160, 130, 145)
        AccentBlue          = @(180, 60, 90)
        TertiaryBg          = @(40, 28, 35)
        TopBarAccent        = @(200, 70, 100)
        InputBgAlt          = @(35, 25, 32)
        SuccessColor        = @(80, 180, 100)
        SuccessColorAlt     = @(80, 180, 100)
        WarningColor        = @(240, 160, 60)
        WarningColorAlt     = @(240, 160, 60)
        ErrorColor          = @(220, 50, 70)
        ErrorColorAlt       = @(220, 50, 70)

        # --- ToolTips ---
        ToolTipBg                   = @(24, 15, 19)
        ToolTipBorder               = @(244, 63, 94)
        ToolTipFg                   = @(250, 238, 242)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(42, 27, 34)
        HeaderButtonBg              = @(30, 20, 24)
        BossButtonBg                = @(68, 20, 30)
        BossTimerText               = @(255, 235, 240)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(22, 14, 18)
        InputFg                     = @(250, 238, 242)
        ComboDropdownBg             = @(30, 20, 24)
        ComboItemText               = @(250, 238, 242)
        ComboBorderNormal           = @(72, 42, 52)
        ComboArrow                  = @(251, 113, 133)
        ComboSelectedItem           = @(190, 24, 65)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(20, 13, 16)
        GridBorder                  = @(48, 30, 36)
        GridSecondaryBg             = @(30, 20, 24)
        GridAlternateRow            = @(24, 16, 20)
        GridSelectedRow             = @(190, 24, 65)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(42, 27, 34)
        GridHeaderText              = @(250, 238, 242)
        GridRowHeaderBg             = @(42, 27, 34)
        GridRowHeaderText           = @(250, 238, 242)

        # --- Tab Controls & Pages ---
        TabBg                       = @(20, 13, 16)
        TabActive                   = @(30, 20, 24)
        TabInactive                 = @(20, 13, 16)
        TabActiveText               = @(250, 238, 242)
        TabInactiveText             = @(175, 138, 148)
        TabUnderline                = @(244, 63, 94)
        TabInactiveUnderline        = @(60, 36, 44)
        TabPageBg                   = @(20, 13, 16)
        TabPageText                 = @(250, 238, 242)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(244, 63, 94)
        ToggleOffColor              = @(65, 30, 38)
        GroupBoxBg                  = @(30, 20, 24)
        GroupBoxText                = @(250, 238, 242)
        ListItemBg                  = @(30, 20, 24)
        ListItemText                = @(250, 238, 242)
        ProgressBarBg               = @(22, 14, 18)
        ProgressBarFill             = @(225, 29, 72)
        TreeViewBg                  = @(22, 14, 18)
        TreeViewHover               = @(46, 28, 36)
        ScrollThumbNormal           = @(60, 36, 44)
        ScrollThumbHover            = @(88, 52, 64)
        ScrollThumbActive           = @(244, 63, 94)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(52, 34, 42)
        BtnGradientNormalBot        = @(36, 22, 28)
        BtnGradientHoverTop         = @(70, 44, 56)
        BtnGradientHoverBot         = @(48, 30, 38)
        BtnGradientLeftPressTop     = @(30, 20, 24)
        BtnGradientLeftPressBot     = @(20, 13, 16)
        BtnGradientRightPressTop    = @(62, 38, 48)
        BtnGradientRightPressBot    = @(42, 26, 34)

        BtnDangerGradientNormalTop    = @(215, 35, 75)
        BtnDangerGradientNormalBot    = @(160, 18, 50)
        BtnDangerGradientHoverTop     = @(240, 52, 95)
        BtnDangerGradientHoverBot     = @(185, 26, 62)
        BtnDangerGradientLeftPressTop = @(140, 14, 42)
        BtnDangerGradientLeftPressBot = @(105, 10, 32)
        BtnDangerGradientRightPressTop= @(192, 38, 80)
        BtnDangerGradientRightPressBot= @(148, 22, 60)

        BtnSuccessGradientNormalTop    = @(48, 148, 96)
        BtnSuccessGradientNormalBot    = @(30, 105, 66)
        BtnSuccessGradientHoverTop     = @(65, 175, 118)
        BtnSuccessGradientHoverBot     = @(42, 132, 84)
        BtnSuccessGradientLeftPressTop = @(26, 92, 58)
        BtnSuccessGradientLeftPressBot = @(18, 68, 42)
        BtnSuccessGradientRightPressTop= @(56, 150, 100)
        BtnSuccessGradientRightPressBot= @(36, 112, 74)

        BtnPrimaryGradientNormalTop    = @(210, 32, 78)
        BtnPrimaryGradientNormalBot    = @(155, 16, 54)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(160, 20, 20)
        StatusDisconnectedText    = @(240, 225, 230)
        StatusFlashing            = @(255, 140, 40)
        StatusFlashingText        = @(40, 15, 20)
        StatusFlashingSelect      = @(220, 110, 20)
        NotificationProgressHigh  = @(210, 60, 100)
        NotificationProgressMid   = @(240, 160, 50)
        NotificationProgressLow   = @(220, 40, 50)
        NotificationStripWarning  = @(230, 100, 40)
        NotificationStripInfo     = @(140, 70, 160)
        NotificationStripError    = @(220, 40, 60)
        NotificationBg            = @(40, 15, 22)
        NotificationStripBg       = @(65, 22, 32)
        NotificationBtnBg         = @(65, 22, 32)
        NotificationBtnHover      = @(90, 32, 46)
        NotificationBtnPressed    = @(30, 10, 16)
        NotificationMsgText       = @(220, 195, 205)
        TooltipBorderColor        = @(120, 40, 60)
        BtnPrimaryGradientHoverTop     = @(238, 52, 102)
        BtnPrimaryGradientHoverBot     = @(180, 24, 68)
        BtnPrimaryGradientLeftPressTop = @(138, 12, 46)
        BtnPrimaryGradientLeftPressBot = @(102, 8, 34)
        BtnPrimaryGradientRightPressTop= @(185, 36, 88)
        BtnPrimaryGradientRightPressBot= @(142, 20, 66)
    }

    # =========================================================================
    # 4. CYBER VIOLET (Dark Mode)
    # =========================================================================
    CyberViolet = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(16, 10, 26)      # Deep void purple
        SecondaryBg                 = @(26, 16, 42)      # Royal indigo layer
        TopBarBg                    = @(11, 7, 18)
        BorderColor                 = @(52, 32, 80)      # Neon purple muted border
        PanelBg                     = @(26, 16, 42)

        # --- Typography & Content ---
        PrimaryText                 = @(246, 238, 255)   # Lavender white
        HeaderTitleText             = @(255, 248, 255)
        HeaderSubtitleText          = @(165, 142, 200)   # Heather violet
        SectionHeaderText           = @(192, 132, 252)   # Radiant lavender
        AccentDanger                = @(251, 113, 133)   # Neon coral
        StatusActive                = @(74, 222, 128)    # Neon lime
        CurrencyPerin               = @(192, 132, 252)   # Violet prism
        CurrencyEcp                 = @(56, 189, 248)    # Cyber cyan

        # --- Semantic Aliases ---
        WhiteText           = @(235, 230, 250)
        SecondaryText       = @(140, 130, 175)
        AccentBlue          = @(130, 80, 220)
        TertiaryBg          = @(30, 25, 50)
        TopBarAccent        = @(150, 100, 240)
        InputBgAlt          = @(22, 18, 42)
        SuccessColor        = @(80, 200, 120)
        SuccessColorAlt     = @(80, 200, 120)
        WarningColor        = @(250, 180, 50)
        WarningColorAlt     = @(250, 180, 50)
        ErrorColor          = @(230, 60, 90)
        ErrorColorAlt       = @(230, 60, 90)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(139, 0, 0)
        StatusDisconnectedText    = @(0, 0, 0)
        StatusFlashing            = @(255, 165, 0)
        StatusFlashingText        = @(0, 0, 0)
        StatusFlashingSelect      = @(255, 140, 0)
        NotificationProgressHigh  = @(168, 120, 247)
        NotificationProgressMid   = @(250, 180, 50)
        NotificationProgressLow   = @(230, 60, 90)
        NotificationStripWarning  = @(230, 140, 50)
        NotificationStripInfo     = @(130, 80, 220)
        NotificationStripError    = @(230, 60, 90)
        NotificationBg            = @(26, 16, 42)
        NotificationStripBg       = @(52, 32, 80)
        NotificationBtnBg         = @(52, 32, 80)
        NotificationBtnHover      = @(78, 48, 118)
        NotificationBtnPressed    = @(38, 24, 62)
        NotificationMsgText       = @(165, 142, 200)
        TooltipBorderColor        = @(168, 85, 247)

        # --- ToolTips ---
        ToolTipBg                   = @(20, 12, 32)
        ToolTipBorder               = @(168, 85, 247)
        ToolTipFg                   = @(246, 238, 255)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(38, 24, 62)
        HeaderButtonBg              = @(26, 16, 42)
        BossButtonBg                = @(65, 22, 45)
        BossTimerText               = @(255, 235, 245)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(18, 11, 30)
        InputFg                     = @(246, 238, 255)
        ComboDropdownBg             = @(26, 16, 42)
        ComboItemText               = @(246, 238, 255)
        ComboBorderNormal           = @(65, 40, 98)
        ComboArrow                  = @(192, 132, 252)
        ComboSelectedItem           = @(147, 51, 234)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(16, 10, 26)
        GridBorder                  = @(40, 24, 62)
        GridSecondaryBg             = @(26, 16, 42)
        GridAlternateRow            = @(20, 12, 32)
        GridSelectedRow             = @(147, 51, 234)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(38, 24, 62)
        GridHeaderText              = @(246, 238, 255)
        GridRowHeaderBg             = @(38, 24, 62)
        GridRowHeaderText           = @(246, 238, 255)

        # --- Tab Controls & Pages ---
        TabBg                       = @(16, 10, 26)
        TabActive                   = @(26, 16, 42)
        TabInactive                 = @(16, 10, 26)
        TabActiveText               = @(246, 238, 255)
        TabInactiveText             = @(165, 142, 200)
        TabUnderline                = @(168, 85, 247)
        TabInactiveUnderline        = @(52, 32, 80)
        TabPageBg                   = @(16, 10, 26)
        TabPageText                 = @(246, 238, 255)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(168, 85, 247)
        ToggleOffColor              = @(65, 25, 52)
        GroupBoxBg                  = @(26, 16, 42)
        GroupBoxText                = @(246, 238, 255)
        ListItemBg                  = @(26, 16, 42)
        ListItemText                = @(246, 238, 255)
        ProgressBarBg               = @(18, 11, 30)
        ProgressBarFill             = @(168, 85, 247)
        TreeViewBg                  = @(18, 11, 30)
        TreeViewHover               = @(42, 25, 68)
        ScrollThumbNormal           = @(52, 32, 80)
        ScrollThumbHover            = @(78, 48, 118)
        ScrollThumbActive           = @(168, 85, 247)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(48, 30, 78)
        BtnGradientNormalBot        = @(32, 20, 52)
        BtnGradientHoverTop         = @(65, 42, 105)
        BtnGradientHoverBot         = @(44, 28, 70)
        BtnGradientLeftPressTop     = @(26, 16, 42)
        BtnGradientLeftPressBot     = @(16, 10, 26)
        BtnGradientRightPressTop    = @(58, 36, 92)
        BtnGradientRightPressBot    = @(38, 24, 62)

        BtnDangerGradientNormalTop    = @(205, 42, 92)
        BtnDangerGradientNormalBot    = @(155, 22, 65)
        BtnDangerGradientHoverTop     = @(230, 58, 115)
        BtnDangerGradientHoverBot     = @(180, 30, 80)
        BtnDangerGradientLeftPressTop = @(135, 18, 55)
        BtnDangerGradientLeftPressBot = @(100, 12, 40)
        BtnDangerGradientRightPressTop= @(188, 45, 100)
        BtnDangerGradientRightPressBot= @(145, 26, 76)

        BtnSuccessGradientNormalTop    = @(45, 148, 98)
        BtnSuccessGradientNormalBot    = @(28, 105, 68)
        BtnSuccessGradientHoverTop     = @(62, 175, 118)
        BtnSuccessGradientHoverBot     = @(40, 132, 85)
        BtnSuccessGradientLeftPressTop = @(25, 92, 58)
        BtnSuccessGradientLeftPressBot = @(18, 68, 42)
        BtnSuccessGradientRightPressTop= @(56, 150, 102)
        BtnSuccessGradientRightPressBot= @(36, 112, 75)

        BtnPrimaryGradientNormalTop    = @(160, 65, 240)
        BtnPrimaryGradientNormalBot    = @(115, 35, 185)
        BtnPrimaryGradientHoverTop     = @(185, 90, 255)
        BtnPrimaryGradientHoverBot     = @(140, 50, 212)
        BtnPrimaryGradientLeftPressTop = @(102, 30, 165)
        BtnPrimaryGradientLeftPressBot = @(75, 20, 125)
        BtnPrimaryGradientRightPressTop= @(150, 60, 220)
        BtnPrimaryGradientRightPressBot= @(110, 40, 170)
    }

    # =========================================================================
    # 5. MIDNIGHT SLATE (Dark Mode)
    # =========================================================================
    MidnightSlate = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(15, 22, 32)      # Deep cool granite
        SecondaryBg                 = @(24, 34, 48)      # Layered slate panel
        TopBarBg                    = @(10, 15, 22)
        BorderColor                 = @(42, 58, 80)      # Slate mist border

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(155, 30, 30)
        StatusDisconnectedText    = @(220, 230, 245)
        StatusFlashing            = @(255, 210, 40)
        StatusFlashingText        = @(15, 22, 32)
        StatusFlashingSelect      = @(210, 175, 20)
        NotificationProgressHigh  = @(120, 210, 60)
        NotificationProgressMid   = @(255, 210, 40)
        NotificationProgressLow   = @(220, 60, 65)
        NotificationStripWarning  = @(255, 200, 40)
        NotificationStripInfo     = @(80, 160, 230)
        NotificationStripError    = @(220, 60, 70)
        NotificationBg            = @(15, 22, 32)
        NotificationStripBg       = @(24, 34, 48)
        NotificationBtnBg         = @(32, 46, 64)
        NotificationBtnHover      = @(42, 58, 80)
        NotificationBtnPressed    = @(10, 15, 22)
        NotificationMsgText       = @(180, 195, 215)
        TooltipBorderColor        = @(52, 72, 100)
        PanelBg                     = @(24, 34, 48)

        # --- Typography & Content ---
        PrimaryText                 = @(241, 245, 249)   # Frosted slate white
        HeaderTitleText             = @(250, 252, 255)
        HeaderSubtitleText          = @(148, 163, 184)   # Blue-gray mist
        SectionHeaderText           = @(163, 230, 53)    # Vivid electric lime
        AccentDanger                = @(248, 113, 113)   # Soft crimson
        StatusActive                = @(163, 230, 53)    # Electric lime active
        CurrencyPerin               = @(56, 189, 248)    # Sky blue
        CurrencyEcp                 = @(163, 230, 53)    # Lime gold

        # --- Semantic Aliases ---
        WhiteText           = @(235, 238, 242)
        SecondaryText       = @(135, 145, 160)
        AccentBlue          = @(70, 130, 200)
        TertiaryBg          = @(32, 36, 48)
        TopBarAccent        = @(80, 140, 210)
        InputBgAlt          = @(24, 28, 38)
        SuccessColor        = @(60, 175, 90)
        SuccessColorAlt     = @(60, 175, 90)
        WarningColor        = @(240, 175, 55)
        WarningColorAlt     = @(240, 175, 55)
        ErrorColor          = @(215, 65, 75)
        ErrorColorAlt       = @(215, 65, 75)

        # --- ToolTips ---
        ToolTipBg                   = @(18, 26, 38)
        ToolTipBorder               = @(163, 230, 53)
        ToolTipFg                   = @(241, 245, 249)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(32, 45, 64)
        HeaderButtonBg              = @(24, 34, 48)
        BossButtonBg                = @(65, 28, 38)
        BossTimerText               = @(255, 235, 240)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(18, 26, 38)
        InputFg                     = @(241, 245, 249)
        ComboDropdownBg             = @(24, 34, 48)
        ComboItemText               = @(241, 245, 249)
        ComboBorderNormal           = @(52, 72, 98)
        ComboArrow                  = @(163, 230, 53)
        ComboSelectedItem           = @(101, 163, 13)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(15, 22, 32)
        GridBorder                  = @(32, 46, 64)
        GridSecondaryBg             = @(24, 34, 48)
        GridAlternateRow            = @(19, 28, 40)
        GridSelectedRow             = @(101, 163, 13)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(32, 45, 64)
        GridHeaderText              = @(241, 245, 249)
        GridRowHeaderBg             = @(32, 45, 64)
        GridRowHeaderText           = @(241, 245, 249)

        # --- Tab Controls & Pages ---
        TabBg                       = @(15, 22, 32)
        TabActive                   = @(24, 34, 48)
        TabInactive                 = @(15, 22, 32)
        TabActiveText               = @(241, 245, 249)
        TabInactiveText             = @(148, 163, 184)
        TabUnderline                = @(163, 230, 53)
        TabInactiveUnderline        = @(42, 58, 80)
        TabPageBg                   = @(15, 22, 32)
        TabPageText                 = @(241, 245, 249)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(163, 230, 53)
        ToggleOffColor              = @(65, 30, 40)
        GroupBoxBg                  = @(24, 34, 48)
        GroupBoxText                = @(241, 245, 249)
        ListItemBg                  = @(24, 34, 48)
        ListItemText                = @(241, 245, 249)
        ProgressBarBg               = @(18, 26, 38)
        ProgressBarFill             = @(132, 204, 22)
        TreeViewBg                  = @(18, 26, 38)
        TreeViewHover               = @(36, 52, 72)
        ScrollThumbNormal           = @(42, 58, 80)
        ScrollThumbHover            = @(62, 85, 116)
        ScrollThumbActive           = @(163, 230, 53)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(40, 56, 78)
        BtnGradientNormalBot        = @(26, 38, 54)
        BtnGradientHoverTop         = @(54, 76, 105)
        BtnGradientHoverBot         = @(36, 50, 70)
        BtnGradientLeftPressTop     = @(24, 34, 48)
        BtnGradientLeftPressBot     = @(15, 22, 32)
        BtnGradientRightPressTop    = @(48, 66, 92)
        BtnGradientRightPressBot    = @(30, 44, 62)

        BtnDangerGradientNormalTop    = @(192, 48, 58)
        BtnDangerGradientNormalBot    = @(142, 28, 38)
        BtnDangerGradientHoverTop     = @(218, 65, 76)
        BtnDangerGradientHoverBot     = @(168, 38, 48)
        BtnDangerGradientLeftPressTop = @(126, 22, 30)
        BtnDangerGradientLeftPressBot = @(95, 14, 22)
        BtnDangerGradientRightPressTop= @(175, 48, 62)
        BtnDangerGradientRightPressBot= @(132, 30, 42)

        BtnSuccessGradientNormalTop    = @(128, 192, 24)
        BtnSuccessGradientNormalBot    = @(90, 140, 14)
        BtnSuccessGradientHoverTop     = @(152, 222, 32)
        BtnSuccessGradientHoverBot     = @(110, 170, 18)
        BtnSuccessGradientLeftPressTop = @(78, 125, 12)
        BtnSuccessGradientLeftPressBot = @(58, 92, 8)
        BtnSuccessGradientRightPressTop= @(115, 180, 28)
        BtnSuccessGradientRightPressBot= @(82, 134, 16)

        BtnPrimaryGradientNormalTop    = @(122, 185, 22)
        BtnPrimaryGradientNormalBot    = @(86, 135, 14)
        BtnPrimaryGradientHoverTop     = @(146, 215, 30)
        BtnPrimaryGradientHoverBot     = @(105, 165, 18)
        BtnPrimaryGradientLeftPressTop = @(75, 120, 10)
        BtnPrimaryGradientLeftPressBot = @(54, 88, 8)
        BtnPrimaryGradientRightPressTop= @(110, 172, 26)
        BtnPrimaryGradientRightPressBot= @(78, 128, 15)
    }

    # =========================================================================
    # 6. ARCTIC FROST (Light Mode)
    # =========================================================================
    ArcticFrost = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(240, 246, 252)   # Powder ice canvas
        SecondaryBg                 = @(255, 255, 255)   # Crisp iceberg white panel
        TopBarBg                    = @(240, 246, 252)
        BorderColor                 = @(202, 222, 240)   # Glacial mist outline
        PanelBg                     = @(255, 255, 255)

        # --- Typography & Content ---
        PrimaryText                 = @(24, 42, 60)      # Deep fjord navy
        HeaderTitleText             = @(14, 28, 44)
        HeaderSubtitleText          = @(104, 134, 162)   # Cool steel mist
        SectionHeaderText           = @(2, 132, 199)     # Deep cerulean
        AccentDanger                = @(225, 29, 72)     # Vivid cranberry
        StatusActive                = @(16, 185, 129)    # Arctic mint
        CurrencyPerin               = @(2, 132, 199)     # Sky blue
        CurrencyEcp                 = @(13, 148, 136)    # Glacial teal

        # --- Semantic Aliases ---
        WhiteText           = @(240, 248, 255)
        SecondaryText       = @(120, 150, 175)
        AccentBlue          = @(60, 160, 220)
        TertiaryBg          = @(28, 45, 65)
        TopBarAccent        = @(70, 170, 230)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(170, 30, 40)
        StatusDisconnectedText    = @(240, 248, 255)
        StatusFlashing            = @(255, 190, 50)
        StatusFlashingText        = @(20, 35, 50)
        StatusFlashingSelect      = @(220, 160, 30)
        NotificationProgressHigh  = @(16, 200, 140)
        NotificationProgressMid   = @(255, 190, 50)
        NotificationProgressLow   = @(225, 50, 70)
        NotificationStripWarning  = @(255, 180, 60)
        NotificationStripInfo     = @(60, 170, 230)
        NotificationStripError    = @(225, 50, 80)
        NotificationBg            = @(20, 35, 55)
        NotificationStripBg       = @(32, 52, 75)
        NotificationBtnBg         = @(36, 58, 82)
        NotificationBtnHover      = @(48, 72, 100)
        NotificationBtnPressed    = @(14, 25, 40)
        NotificationMsgText       = @(210, 225, 240)
        TooltipBorderColor        = @(60, 100, 145)
        InputBgAlt          = @(20, 35, 55)
        SuccessColor        = @(55, 190, 130)
        SuccessColorAlt     = @(55, 190, 130)
        WarningColor        = @(245, 190, 60)
        WarningColorAlt     = @(245, 190, 60)
        ErrorColor          = @(220, 80, 95)
        ErrorColorAlt       = @(220, 80, 95)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(180, 212, 238)
        ToolTipFg                   = @(24, 42, 60)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(230, 240, 250)
        HeaderButtonBg              = @(240, 246, 252)
        BossButtonBg                = @(254, 228, 232)
        BossTimerText               = @(24, 42, 60)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(24, 42, 60)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(24, 42, 60)
        ComboBorderNormal           = @(190, 216, 238)
        ComboArrow                  = @(2, 132, 199)
        ComboSelectedItem           = @(2, 132, 199)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(240, 246, 252)
        GridBorder                  = @(220, 234, 246)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(245, 250, 254)
        GridSelectedRow             = @(2, 132, 199)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(226, 238, 249)
        GridHeaderText              = @(24, 42, 60)
        GridRowHeaderBg             = @(226, 238, 249)
        GridRowHeaderText           = @(24, 42, 60)

        # --- Tab Controls & Pages ---
        TabBg                       = @(240, 246, 252)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(240, 246, 252)
        TabActiveText               = @(24, 42, 60)
        TabInactiveText             = @(104, 134, 162)
        TabUnderline                = @(2, 132, 199)
        TabInactiveUnderline        = @(202, 222, 240)
        TabPageBg                   = @(240, 246, 252)
        TabPageText                 = @(24, 42, 60)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(2, 132, 199)
        ToggleOffColor              = @(220, 75, 75)
        GroupBoxBg                  = @(255, 255, 255)
        GroupBoxText                = @(24, 42, 60)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(24, 42, 60)
        ProgressBarBg               = @(226, 238, 249)
        ProgressBarFill             = @(2, 132, 199)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(232, 244, 254)
        ScrollThumbNormal           = @(192, 216, 236)
        ScrollThumbHover            = @(158, 192, 220)
        ScrollThumbActive           = @(2, 132, 199)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(250, 253, 255)
        BtnGradientNormalBot        = @(226, 238, 249)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(236, 246, 254)
        BtnGradientLeftPressTop     = @(216, 232, 246)
        BtnGradientLeftPressBot     = @(198, 218, 238)
        BtnGradientRightPressTop    = @(225, 238, 250)
        BtnGradientRightPressBot    = @(208, 226, 242)

        BtnDangerGradientNormalTop    = @(252, 130, 130)
        BtnDangerGradientNormalBot    = @(225, 29, 72)
        BtnDangerGradientHoverTop     = @(254, 155, 155)
        BtnDangerGradientHoverBot     = @(240, 52, 92)
        BtnDangerGradientLeftPressTop = @(205, 20, 60)
        BtnDangerGradientLeftPressBot = @(175, 12, 45)
        BtnDangerGradientRightPressTop= @(232, 95, 120)
        BtnDangerGradientRightPressBot= @(200, 42, 75)

        BtnSuccessGradientNormalTop    = @(102, 224, 172)
        BtnSuccessGradientNormalBot    = @(16, 185, 129)
        BtnSuccessGradientHoverTop     = @(134, 240, 192)
        BtnSuccessGradientHoverBot     = @(32, 202, 142)
        BtnSuccessGradientLeftPressTop = @(10, 160, 110)
        BtnSuccessGradientLeftPressBot = @(6, 130, 88)
        BtnSuccessGradientRightPressTop= @(78, 204, 156)
        BtnSuccessGradientRightPressBot= @(24, 166, 120)

        BtnPrimaryGradientNormalTop    = @(56, 189, 248)
        BtnPrimaryGradientNormalBot    = @(2, 132, 199)
        BtnPrimaryGradientHoverTop     = @(92, 208, 254)
        BtnPrimaryGradientHoverBot     = @(14, 152, 222)
        BtnPrimaryGradientLeftPressTop = @(2, 115, 175)
        BtnPrimaryGradientLeftPressBot = @(1, 90, 140)
        BtnPrimaryGradientRightPressTop= @(40, 162, 232)
        BtnPrimaryGradientRightPressBot= @(8, 122, 186)
    }

    # =========================================================================
    # 7. SUNNY MEADOW (Light Mode)
    # =========================================================================
    SunnyMeadow = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(252, 250, 240)   # Soft buttermilk canvas
        SecondaryBg                 = @(255, 255, 255)   # Warm white panel
        TopBarBg                    = @(252, 250, 240)
        BorderColor                 = @(232, 224, 200)   # Golden straw border
        PanelBg                     = @(255, 255, 255)

        # --- Typography & Content ---
        PrimaryText                 = @(48, 42, 32)      # Deep earth brown
        HeaderTitleText             = @(32, 26, 18)
        HeaderSubtitleText          = @(145, 135, 115)   # Warm flaxen gray
        SectionHeaderText           = @(101, 163, 13)    # Fresh clover green
        AccentDanger                = @(220, 53, 69)     # Poppy red
        StatusActive                = @(22, 163, 74)     # Leaf emerald
        CurrencyPerin               = @(217, 119, 6)     # Amber honey
        CurrencyEcp                 = @(101, 163, 13)    # Meadow green

        # --- Semantic Aliases ---
        WhiteText           = @(250, 248, 240)
        SecondaryText       = @(140, 135, 115)
        AccentBlue          = @(100, 150, 60)
        TertiaryBg          = @(42, 45, 32)
        TopBarAccent        = @(120, 170, 70)
        InputBgAlt          = @(35, 38, 28)
        SuccessColor        = @(80, 170, 80)
        SuccessColorAlt     = @(80, 170, 80)
        WarningColor        = @(230, 170, 40)
        WarningColorAlt     = @(230, 170, 40)
        ErrorColor          = @(200, 70, 60)
        ErrorColorAlt       = @(200, 70, 60)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(217, 119, 6)
        ToolTipFg                   = @(48, 42, 32)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(246, 241, 224)
        HeaderButtonBg              = @(252, 250, 240)
        BossButtonBg                = @(254, 232, 224)
        BossTimerText               = @(48, 42, 32)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(180, 40, 30)
        StatusDisconnectedText    = @(255, 255, 255)
        StatusFlashing            = @(217, 119, 6)
        StatusFlashingText        = @(48, 42, 32)
        StatusFlashingSelect      = @(180, 95, 0)
        NotificationProgressHigh  = @(40, 160, 55)
        NotificationProgressMid   = @(220, 160, 30)
        NotificationProgressLow   = @(200, 60, 45)
        NotificationStripWarning  = @(217, 119, 6)
        NotificationStripInfo     = @(60, 150, 100)
        NotificationStripError    = @(200, 60, 50)
        NotificationBg            = @(255, 252, 242)
        NotificationStripBg       = @(246, 241, 224)
        NotificationBtnBg         = @(240, 232, 210)
        NotificationBtnHover      = @(224, 216, 198)
        NotificationBtnPressed    = @(210, 200, 180)
        NotificationMsgText       = @(68, 62, 48)
        TooltipBorderColor        = @(217, 190, 140)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(48, 42, 32)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(48, 42, 32)
        ComboBorderNormal           = @(225, 215, 190)
        ComboArrow                  = @(101, 163, 13)
        ComboSelectedItem           = @(101, 163, 13)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(252, 250, 240)
        GridBorder                  = @(238, 230, 212)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(248, 246, 235)
        GridSelectedRow             = @(101, 163, 13)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(242, 236, 215)
        GridHeaderText              = @(48, 42, 32)
        GridRowHeaderBg             = @(242, 236, 215)
        GridRowHeaderText           = @(48, 42, 32)

        # --- Tab Controls & Pages ---
        TabBg                       = @(252, 250, 240)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(252, 250, 240)
        TabActiveText               = @(48, 42, 32)
        TabInactiveText             = @(145, 135, 115)
        TabUnderline                = @(101, 163, 13)
        TabInactiveUnderline        = @(232, 224, 200)
        TabPageBg                   = @(252, 250, 240)
        TabPageText                 = @(48, 42, 32)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(101, 163, 13)
        ToggleOffColor              = @(205, 60, 60)
        GroupBoxBg                  = @(255, 255, 255)
        GroupBoxText                = @(48, 42, 32)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(48, 42, 32)
        ProgressBarBg               = @(242, 236, 215)
        ProgressBarFill             = @(101, 163, 13)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(245, 240, 220)
        ScrollThumbNormal           = @(220, 210, 185)
        ScrollThumbHover            = @(192, 180, 150)
        ScrollThumbActive           = @(101, 163, 13)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(254, 252, 245)
        BtnGradientNormalBot        = @(240, 234, 212)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(246, 242, 224)
        BtnGradientLeftPressTop     = @(226, 218, 192)
        BtnGradientLeftPressBot     = @(212, 202, 175)
        BtnGradientRightPressTop    = @(236, 228, 204)
        BtnGradientRightPressBot    = @(218, 208, 182)

        BtnDangerGradientNormalTop    = @(245, 125, 110)
        BtnDangerGradientNormalBot    = @(220, 53, 69)
        BtnDangerGradientHoverTop     = @(252, 145, 130)
        BtnDangerGradientHoverBot     = @(232, 70, 85)
        BtnDangerGradientLeftPressTop = @(200, 42, 56)
        BtnDangerGradientLeftPressBot = @(170, 28, 40)
        BtnDangerGradientRightPressTop= @(230, 100, 95)
        BtnDangerGradientRightPressBot= @(198, 48, 62)

        BtnSuccessGradientNormalTop    = @(148, 214, 52)
        BtnSuccessGradientNormalBot    = @(101, 163, 13)
        BtnSuccessGradientHoverTop     = @(168, 230, 72)
        BtnSuccessGradientHoverBot     = @(118, 185, 20)
        BtnSuccessGradientLeftPressTop = @(88, 142, 10)
        BtnSuccessGradientLeftPressBot = @(68, 112, 8)
        BtnSuccessGradientRightPressTop= @(128, 194, 42)
        BtnSuccessGradientRightPressBot= @(90, 148, 14)

        BtnPrimaryGradientNormalTop    = @(245, 160, 42)
        BtnPrimaryGradientNormalBot    = @(217, 119, 6)
        BtnPrimaryGradientHoverTop     = @(252, 180, 68)
        BtnPrimaryGradientHoverBot     = @(234, 134, 14)
        BtnPrimaryGradientLeftPressTop = @(194, 104, 5)
        BtnPrimaryGradientLeftPressBot = @(164, 86, 4)
        BtnPrimaryGradientRightPressTop= @(224, 138, 32)
        BtnPrimaryGradientRightPressBot= @(184, 96, 8)
    }

    # =========================================================================
    # 8. QUARTZ STONE (Light Mode)
    # =========================================================================
    QuartzStone = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(245, 245, 247)   # Neutral alabaster
        SecondaryBg                 = @(255, 255, 255)   # Crisp white slab
        TopBarBg                    = @(245, 245, 247)
        BorderColor                 = @(220, 220, 226)   # Silvery stone outline
        PanelBg                     = @(255, 255, 255)

        # --- Typography & Content ---
        PrimaryText                 = @(38, 38, 44)      # Architectural charcoal
        HeaderTitleText             = @(20, 20, 26)
        HeaderSubtitleText          = @(130, 130, 140)   # Neutral slate mist
        SectionHeaderText           = @(244, 114, 182)   # Rose quartz blossom
        AccentDanger                = @(244, 63, 94)     # Crisp rubellite
        StatusActive                = @(34, 197, 94)     # Jade green
        CurrencyPerin               = @(148, 163, 184)   # Slate platinum
        CurrencyEcp                 = @(244, 114, 182)   # Rose quartz

        # --- Semantic Aliases ---
        WhiteText           = @(240, 238, 235)
        SecondaryText       = @(145, 140, 135)
        AccentBlue          = @(120, 110, 150)
        TertiaryBg          = @(48, 45, 42)
        TopBarAccent        = @(140, 130, 170)
        InputBgAlt          = @(42, 40, 38)
        SuccessColor        = @(90, 160, 100)
        SuccessColorAlt     = @(90, 160, 100)
        WarningColor        = @(220, 165, 55)
        WarningColorAlt     = @(220, 165, 55)
        ErrorColor          = @(195, 75, 70)
        ErrorColorAlt       = @(195, 75, 70)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(244, 114, 182)
        ToolTipFg                   = @(38, 38, 44)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(236, 236, 240)
        HeaderButtonBg              = @(245, 245, 247)
        BossButtonBg                = @(254, 228, 234)
        BossTimerText               = @(38, 38, 44)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(165, 45, 45)
        StatusDisconnectedText    = @(248, 248, 250)
        StatusFlashing            = @(224, 160, 50)
        StatusFlashingText        = @(42, 38, 36)
        StatusFlashingSelect      = @(190, 130, 30)
        NotificationProgressHigh  = @(50, 170, 80)
        NotificationProgressMid   = @(224, 170, 50)
        NotificationProgressLow   = @(200, 60, 60)
        NotificationStripWarning  = @(210, 140, 40)
        NotificationStripInfo     = @(80, 120, 180)
        NotificationStripError    = @(195, 65, 65)
        NotificationBg            = @(244, 244, 248)
        NotificationStripBg       = @(224, 224, 230)
        NotificationBtnBg         = @(218, 218, 224)
        NotificationBtnHover      = @(200, 200, 208)
        NotificationBtnPressed    = @(190, 190, 196)
        NotificationMsgText       = @(55, 52, 58)
        TooltipBorderColor        = @(195, 195, 202)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(38, 38, 44)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(38, 38, 44)
        ComboBorderNormal           = @(218, 218, 224)
        ComboArrow                  = @(244, 114, 182)
        ComboSelectedItem           = @(244, 114, 182)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(245, 245, 247)
        GridBorder                  = @(228, 228, 234)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(240, 240, 243)
        GridSelectedRow             = @(244, 114, 182)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(232, 232, 238)
        GridHeaderText              = @(38, 38, 44)
        GridRowHeaderBg             = @(232, 232, 238)
        GridRowHeaderText           = @(38, 38, 44)

        # --- Tab Controls & Pages ---
        TabBg                       = @(245, 245, 247)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(245, 245, 247)
        TabActiveText               = @(38, 38, 44)
        TabInactiveText             = @(130, 130, 140)
        TabUnderline                = @(244, 114, 182)
        TabInactiveUnderline        = @(220, 220, 226)
        TabPageBg                   = @(245, 245, 247)
        TabPageText                 = @(38, 38, 44)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(244, 114, 182)
        ToggleOffColor              = @(185, 185, 195)
        GroupBoxBg                  = @(255, 255, 255)
        GroupBoxText                = @(38, 38, 44)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(38, 38, 44)
        ProgressBarBg               = @(236, 236, 240)
        ProgressBarFill             = @(244, 114, 182)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(242, 238, 244)
        ScrollThumbNormal           = @(208, 208, 216)
        ScrollThumbHover            = @(178, 178, 188)
        ScrollThumbActive           = @(244, 114, 182)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(250, 250, 252)
        BtnGradientNormalBot        = @(232, 232, 238)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(238, 238, 244)
        BtnGradientLeftPressTop     = @(218, 218, 224)
        BtnGradientLeftPressBot     = @(202, 202, 210)
        BtnGradientRightPressTop    = @(228, 225, 234)
        BtnGradientRightPressBot    = @(210, 206, 218)

        BtnDangerGradientNormalTop    = @(250, 128, 150)
        BtnDangerGradientNormalBot    = @(244, 63, 94)
        BtnDangerGradientHoverTop     = @(254, 152, 170)
        BtnDangerGradientHoverBot     = @(248, 85, 112)
        BtnDangerGradientLeftPressTop = @(224, 45, 75)
        BtnDangerGradientLeftPressBot = @(194, 30, 58)
        BtnDangerGradientRightPressTop= @(238, 104, 130)
        BtnDangerGradientRightPressBot= @(206, 54, 82)

        BtnSuccessGradientNormalTop    = @(118, 224, 158)
        BtnSuccessGradientNormalBot    = @(34, 197, 94)
        BtnSuccessGradientHoverTop     = @(146, 238, 180)
        BtnSuccessGradientHoverBot     = @(54, 212, 112)
        BtnSuccessGradientLeftPressTop = @(24, 170, 80)
        BtnSuccessGradientLeftPressBot = @(18, 140, 65)
        BtnSuccessGradientRightPressTop= @(90, 204, 140)
        BtnSuccessGradientRightPressBot= @(40, 164, 88)

        BtnPrimaryGradientNormalTop    = @(248, 142, 196)
        BtnPrimaryGradientNormalBot    = @(236, 72, 153)
        BtnPrimaryGradientHoverTop     = @(252, 168, 210)
        BtnPrimaryGradientHoverBot     = @(242, 94, 168)
        BtnPrimaryGradientLeftPressTop = @(214, 50, 130)
        BtnPrimaryGradientLeftPressBot = @(184, 34, 108)
        BtnPrimaryGradientRightPressTop= @(230, 108, 180)
        BtnPrimaryGradientRightPressBot= @(196, 64, 140)
    }

    # =========================================================================
    # 9. TERRA COTTA (Warm Light Mode)
    # =========================================================================
    TerraCotta = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(250, 242, 236)   # Desert sand canvas
        SecondaryBg                 = @(255, 255, 255)   # Cream white panel
        TopBarBg                    = @(250, 242, 236)
        BorderColor                 = @(234, 212, 200)   # Baked clay border
        PanelBg                     = @(255, 255, 255)

        # --- Typography & Content ---
        PrimaryText                 = @(54, 36, 28)      # Deep espresso earth
        HeaderTitleText             = @(36, 20, 14)
        HeaderSubtitleText          = @(152, 120, 106)   # Russet clay mist
        SectionHeaderText           = @(194, 65, 12)     # Earthen burnt orange
        AccentDanger                = @(220, 38, 38)     # Carmine clay
        StatusActive                = @(22, 163, 74)     # Warm green sprig
        CurrencyPerin               = @(194, 65, 12)     # Terracotta
        CurrencyEcp                 = @(217, 119, 6)     # Warm ochre

        # --- Semantic Aliases ---
        WhiteText           = @(245, 238, 230)
        SecondaryText       = @(155, 135, 120)
        AccentBlue          = @(180, 100, 60)
        TertiaryBg          = @(50, 38, 30)
        TopBarAccent        = @(200, 120, 70)
        InputBgAlt          = @(42, 32, 25)
        SuccessColor        = @(80, 165, 85)
        SuccessColorAlt     = @(80, 165, 85)
        WarningColor        = @(230, 165, 45)
        WarningColorAlt     = @(230, 165, 45)
        ErrorColor          = @(210, 65, 55)
        ErrorColorAlt       = @(210, 65, 55)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(194, 65, 12)
        ToolTipFg                   = @(54, 36, 28)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(244, 228, 218)
        HeaderButtonBg              = @(250, 242, 236)
        BossButtonBg                = @(254, 222, 214)
        BossTimerText               = @(54, 36, 28)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(54, 36, 28)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(54, 36, 28)
        ComboBorderNormal           = @(228, 206, 194)
        ComboArrow                  = @(194, 65, 12)
        ComboSelectedItem           = @(194, 65, 12)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(250, 242, 236)
        GridBorder                  = @(238, 220, 210)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(246, 236, 228)
        GridSelectedRow             = @(194, 65, 12)
        GridSelectedText            = @(255, 255, 255)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(170, 50, 30)
        StatusDisconnectedText    = @(250, 240, 234)
        StatusFlashing            = @(230, 150, 40)
        StatusFlashingText        = @(54, 36, 28)
        StatusFlashingSelect      = @(195, 120, 20)
        NotificationProgressHigh  = @(60, 165, 65)
        NotificationProgressMid   = @(230, 160, 40)
        NotificationProgressLow   = @(210, 55, 45)
        NotificationStripWarning  = @(220, 130, 30)
        NotificationStripInfo     = @(80, 130, 160)
        NotificationStripError    = @(200, 65, 50)
        NotificationBg            = @(250, 242, 236)
        NotificationStripBg       = @(235, 222, 210)
        NotificationBtnBg         = @(228, 214, 200)
        NotificationBtnHover      = @(212, 198, 184)
        NotificationBtnPressed    = @(200, 185, 170)
        NotificationMsgText       = @(72, 55, 44)
        TooltipBorderColor        = @(200, 175, 155)
        GridHeaderBg                = @(240, 224, 212)
        GridHeaderText              = @(54, 36, 28)
        GridRowHeaderBg             = @(240, 224, 212)
        GridRowHeaderText           = @(54, 36, 28)

        # --- Tab Controls & Pages ---
        TabBg                       = @(250, 242, 236)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(250, 242, 236)
        TabActiveText               = @(54, 36, 28)
        TabInactiveText             = @(152, 120, 106)
        TabUnderline                = @(194, 65, 12)
        TabInactiveUnderline        = @(234, 212, 200)
        TabPageBg                   = @(250, 242, 236)
        TabPageText                 = @(54, 36, 28)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(194, 65, 12)
        ToggleOffColor              = @(185, 75, 75)
        GroupBoxBg                  = @(255, 255, 255)
        GroupBoxText                = @(54, 36, 28)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(54, 36, 28)
        ProgressBarBg               = @(244, 228, 218)
        ProgressBarFill             = @(194, 65, 12)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(245, 230, 220)
        ScrollThumbNormal           = @(224, 200, 186)
        ScrollThumbHover            = @(198, 168, 150)
        ScrollThumbActive           = @(194, 65, 12)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(252, 246, 240)
        BtnGradientNormalBot        = @(238, 222, 210)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(244, 230, 220)
        BtnGradientLeftPressTop     = @(224, 204, 190)
        BtnGradientLeftPressBot     = @(206, 184, 170)
        BtnGradientRightPressTop    = @(232, 214, 202)
        BtnGradientRightPressBot    = @(214, 192, 180)

        BtnDangerGradientNormalTop    = @(245, 115, 115)
        BtnDangerGradientNormalBot    = @(220, 38, 38)
        BtnDangerGradientHoverTop     = @(252, 140, 140)
        BtnDangerGradientHoverBot     = @(235, 55, 55)
        BtnDangerGradientLeftPressTop = @(200, 28, 28)
        BtnDangerGradientLeftPressBot = @(170, 18, 18)
        BtnDangerGradientRightPressTop= @(228, 90, 90)
        BtnDangerGradientRightPressBot= @(195, 45, 45)

        BtnSuccessGradientNormalTop    = @(118, 214, 134)
        BtnSuccessGradientNormalBot    = @(22, 163, 74)
        BtnSuccessGradientHoverTop     = @(144, 230, 158)
        BtnSuccessGradientHoverBot     = @(34, 184, 88)
        BtnSuccessGradientLeftPressTop = @(18, 142, 62)
        BtnSuccessGradientLeftPressBot = @(12, 114, 48)
        BtnSuccessGradientRightPressTop= @(84, 194, 118)
        BtnSuccessGradientRightPressBot= @(30, 150, 74)

        BtnPrimaryGradientNormalTop    = @(242, 122, 58)
        BtnPrimaryGradientNormalBot    = @(194, 65, 12)
        BtnPrimaryGradientHoverTop     = @(252, 145, 82)
        BtnPrimaryGradientHoverBot     = @(214, 80, 22)
        BtnPrimaryGradientLeftPressTop = @(174, 54, 10)
        BtnPrimaryGradientLeftPressBot = @(144, 40, 6)
        BtnPrimaryGradientRightPressTop= @(216, 98, 44)
        BtnPrimaryGradientRightPressBot= @(176, 66, 18)
    }

    # =========================================================================
    # 10. DEEP OCEAN (Light Mode)
    # =========================================================================
    DeepOceanLight = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(242, 246, 252)   # Maritime mist canvas
        SecondaryBg                 = @(255, 255, 255)   # Clean white bridge deck
        TopBarBg                    = @(242, 246, 252)
        BorderColor                 = @(204, 218, 236)   # Nautical blue-gray outline
        PanelBg                     = @(255, 255, 255)

        # --- Typography & Content ---
        PrimaryText                 = @(24, 42, 66)      # Naval officer navy
        HeaderTitleText             = @(14, 26, 46)
        HeaderSubtitleText          = @(98, 130, 168)    # Foggy harbor blue
        SectionHeaderText           = @(14, 116, 204)    # Vivid ocean azure
        AccentDanger                = @(225, 29, 72)     # Beacon red
        StatusActive                = @(16, 185, 129)    # Starboard green
        CurrencyPerin               = @(14, 116, 204)    # Ocean blue
        CurrencyEcp                 = @(37, 99, 235)     # Deep royal sapphire

        # --- Semantic Aliases ---
        WhiteText           = @(255, 255, 255)
        SecondaryText       = @(100, 120, 145)
        AccentBlue          = @(14, 116, 204)
        TertiaryBg          = @(230, 238, 248)
        TopBarAccent        = @(14, 116, 204)
        InputBgAlt          = @(240, 246, 252)
        SuccessColor        = @(16, 150, 100)
        SuccessColorAlt     = @(16, 150, 100)
        WarningColor        = @(190, 140, 20)
        WarningColorAlt     = @(190, 140, 20)
        ErrorColor          = @(190, 45, 55)
        ErrorColorAlt       = @(190, 45, 55)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(14, 116, 204)
        ToolTipFg                   = @(24, 42, 66)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(232, 240, 250)
        HeaderButtonBg              = @(242, 246, 252)
        BossButtonBg                = @(254, 226, 228)
        BossTimerText               = @(24, 42, 66)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(24, 42, 66)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(24, 42, 66)
        ComboBorderNormal           = @(198, 214, 234)
        ComboArrow                  = @(14, 116, 204)
        ComboSelectedItem           = @(14, 116, 204)
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(242, 246, 252)
        GridBorder                  = @(220, 232, 246)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(246, 250, 254)
        GridSelectedRow             = @(14, 116, 204)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(226, 236, 248)
        GridHeaderText              = @(24, 42, 66)
        GridRowHeaderBg             = @(226, 236, 248)
        GridRowHeaderText           = @(24, 42, 66)

        # --- Tab Controls & Pages ---
        TabBg                       = @(242, 246, 252)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(242, 246, 252)
        TabActiveText               = @(24, 42, 66)
        TabInactiveText             = @(98, 130, 168)
        TabUnderline                = @(14, 116, 204)
        TabInactiveUnderline        = @(204, 218, 236)
        TabPageBg                   = @(242, 246, 252)
        TabPageText                 = @(24, 42, 66)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(14, 116, 204)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(165, 30, 30)
        StatusDisconnectedText    = @(240, 246, 252)
        StatusFlashing            = @(255, 180, 40)
        StatusFlashingText        = @(24, 42, 66)
        StatusFlashingSelect      = @(210, 145, 20)
        NotificationProgressHigh  = @(14, 160, 140)
        NotificationProgressMid   = @(255, 180, 40)
        NotificationProgressLow   = @(210, 55, 65)
        NotificationStripWarning  = @(0, 160, 200)
        NotificationStripInfo     = @(56, 140, 210)
        NotificationStripError    = @(200, 55, 70)
        NotificationBg            = @(232, 240, 250)
        NotificationStripBg       = @(220, 232, 246)
        NotificationBtnBg         = @(210, 224, 240)
        NotificationBtnHover      = @(195, 212, 232)
        NotificationBtnPressed    = @(185, 200, 220)
        NotificationMsgText       = @(36, 56, 80)
        TooltipBorderColor        = @(160, 190, 220)
        ToggleOffColor              = @(215, 65, 65)
        GroupBoxBg                  = @(255, 255, 255)
        GroupBoxText                = @(24, 42, 66)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(24, 42, 66)
        ProgressBarBg               = @(232, 240, 250)
        ProgressBarFill             = @(14, 116, 204)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(230, 240, 252)
        ScrollThumbNormal           = @(196, 214, 234)
        ScrollThumbHover            = @(162, 188, 218)
        ScrollThumbActive           = @(14, 116, 204)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(248, 252, 255)
        BtnGradientNormalBot        = @(228, 238, 248)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(236, 244, 254)
        BtnGradientLeftPressTop     = @(216, 228, 244)
        BtnGradientLeftPressBot     = @(198, 214, 234)
        BtnGradientRightPressTop    = @(224, 234, 248)
        BtnGradientRightPressBot    = @(206, 222, 240)

        BtnDangerGradientNormalTop    = @(248, 128, 128)
        BtnDangerGradientNormalBot    = @(225, 29, 72)
        BtnDangerGradientHoverTop     = @(254, 152, 152)
        BtnDangerGradientHoverBot     = @(238, 54, 94)
        BtnDangerGradientLeftPressTop = @(204, 20, 60)
        BtnDangerGradientLeftPressBot = @(174, 12, 45)
        BtnDangerGradientRightPressTop= @(232, 90, 115)
        BtnDangerGradientRightPressBot= @(200, 40, 75)

        BtnSuccessGradientNormalTop    = @(110, 228, 175)
        BtnSuccessGradientNormalBot    = @(16, 185, 129)
        BtnSuccessGradientHoverTop     = @(140, 242, 195)
        BtnSuccessGradientHoverBot     = @(35, 205, 145)
        BtnSuccessGradientLeftPressTop = @(10, 160, 110)
        BtnSuccessGradientLeftPressBot = @(6, 130, 88)
        BtnSuccessGradientRightPressTop= @(80, 208, 158)
        BtnSuccessGradientRightPressBot= @(25, 168, 122)

        BtnPrimaryGradientNormalTop    = @(56, 160, 248)
        BtnPrimaryGradientNormalBot    = @(14, 116, 204)
        BtnPrimaryGradientHoverTop     = @(88, 180, 255)
        BtnPrimaryGradientHoverBot     = @(28, 135, 225)
        BtnPrimaryGradientLeftPressTop = @(10, 98, 178)
        BtnPrimaryGradientLeftPressBot = @(6, 75, 142)
        BtnPrimaryGradientRightPressTop= @(40, 140, 225)
        BtnPrimaryGradientRightPressBot= @(18, 105, 185)
    }

    # =========================================================================
    # 11. FLYFF CLASSIC (Light Mode)
    # =========================================================================
    FlyffClassic = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(237, 241, 246)   # Original Flyff window body (cool white)
        SecondaryBg                 = @(255, 255, 255)   # Crisp white content panels
        TopBarBg                    = @(210, 224, 240)   # Pale blue title bar gradient bottom stop
        BorderColor                 = @(174, 186, 202)   # Muted blue-gray outlines
        PanelBg                     = @(250, 251, 253)

        # --- Typography & Content ---
        PrimaryText                 = @(35, 42, 54)      # Near-black client text
        HeaderTitleText             = @(32, 60, 105)     # Deep Flyff blue title
        HeaderSubtitleText          = @(105, 120, 140)   # Faded slate subtitle
        SectionHeaderText           = @(60, 120, 190)    # Flyff sky blue
        AccentDanger                = @(200, 55, 55)
        StatusActive                = @(40, 150, 70)
        CurrencyPerin               = @(200, 160, 40)    # Gold perin coin
        CurrencyEcp                 = @(60, 150, 70)     # Green ecoin

        # --- Semantic Aliases ---
        WhiteText           = @(255, 255, 255)
        SecondaryText       = @(105, 120, 140)
        AccentBlue          = @(74, 130, 200)
        TertiaryBg          = @(230, 236, 245)
        TopBarAccent        = @(90, 150, 215)
        InputBgAlt          = @(255, 255, 255)
        SuccessColor        = @(40, 150, 70)
        SuccessColorAlt     = @(40, 150, 70)
        WarningColor        = @(200, 150, 30)
        WarningColorAlt     = @(200, 150, 30)
        ErrorColor          = @(200, 55, 55)
        ErrorColorAlt       = @(200, 55, 55)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(180, 40, 40)
        StatusDisconnectedText    = @(255, 255, 255)
        StatusFlashing            = @(255, 150, 20)
        StatusFlashingText        = @(0, 0, 0)
        StatusFlashingSelect      = @(220, 130, 10)
        NotificationProgressHigh  = @(60, 175, 80)
        NotificationProgressMid   = @(230, 175, 40)
        NotificationProgressLow   = @(210, 65, 65)
        NotificationStripWarning  = @(255, 150, 20)
        NotificationStripInfo     = @(70, 130, 210)
        NotificationStripError    = @(210, 75, 75)
        NotificationBg            = @(255, 255, 255)
        NotificationStripBg       = @(232, 237, 244)
        NotificationBtnBg         = @(225, 231, 240)
        NotificationBtnHover      = @(205, 215, 230)
        NotificationBtnPressed    = @(190, 202, 220)
        NotificationMsgText       = @(55, 62, 75)
        TooltipBorderColor        = @(174, 186, 202)

        # --- ToolTips ---
        ToolTipBg                   = @(255, 255, 255)
        ToolTipBorder               = @(150, 165, 185)
        ToolTipFg                   = @(35, 42, 54)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(225, 231, 240)   # Classic light blue-gray button
        HeaderButtonBg              = @(232, 240, 250)   # Title bar gradient top stop
        BossButtonBg                = @(250, 220, 215)
        BossTimerText               = @(45, 45, 45)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(255, 255, 255)
        InputFg                     = @(35, 42, 54)
        ComboDropdownBg             = @(255, 255, 255)
        ComboItemText               = @(35, 42, 54)
        ComboBorderNormal           = @(174, 186, 202)
        ComboArrow                  = @(74, 130, 200)
        ComboSelectedItem           = @(49, 106, 197)    # Classic XP selection blue
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(248, 250, 253)
        GridBorder                  = @(214, 222, 234)
        GridSecondaryBg             = @(255, 255, 255)
        GridAlternateRow            = @(243, 246, 251)
        GridSelectedRow             = @(49, 106, 197)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(228, 234, 244)
        GridHeaderText              = @(35, 42, 54)
        GridRowHeaderBg             = @(228, 234, 244)
        GridRowHeaderText           = @(35, 42, 54)

        # --- Tab Controls & Pages ---
        TabBg                       = @(237, 241, 246)
        TabActive                   = @(255, 255, 255)
        TabInactive                 = @(230, 236, 245)
        TabActiveText               = @(35, 42, 54)
        TabInactiveText             = @(125, 135, 150)
        TabUnderline                = @(74, 130, 200)
        TabInactiveUnderline        = @(214, 222, 234)
        TabPageBg                   = @(237, 241, 246)
        TabPageText                 = @(35, 42, 54)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(45, 170, 80)
        ToggleOffColor              = @(205, 70, 70)
        GroupBoxBg                  = @(250, 251, 253)
        GroupBoxText                = @(35, 42, 54)
        ListItemBg                  = @(255, 255, 255)
        ListItemText                = @(35, 42, 54)
        ProgressBarBg               = @(224, 230, 240)
        ProgressBarFill             = @(74, 150, 210)
        TreeViewBg                  = @(255, 255, 255)
        TreeViewHover               = @(224, 234, 248)
        ScrollThumbNormal           = @(200, 208, 220)
        ScrollThumbHover            = @(178, 190, 208)
        ScrollThumbActive           = @(74, 130, 200)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(250, 252, 254)
        BtnGradientNormalBot        = @(230, 236, 245)
        BtnGradientHoverTop         = @(255, 255, 255)
        BtnGradientHoverBot         = @(237, 244, 253)
        BtnGradientLeftPressTop     = @(216, 224, 238)
        BtnGradientLeftPressBot     = @(200, 210, 228)
        BtnGradientRightPressTop    = @(224, 234, 248)
        BtnGradientRightPressBot    = @(206, 219, 240)

        BtnDangerGradientNormalTop    = @(235, 130, 130)
        BtnDangerGradientNormalBot    = @(210, 95, 95)
        BtnDangerGradientHoverTop     = @(245, 155, 155)
        BtnDangerGradientHoverBot     = @(225, 115, 115)
        BtnDangerGradientLeftPressTop = @(205, 90, 90)
        BtnDangerGradientLeftPressBot = @(185, 70, 70)
        BtnDangerGradientRightPressTop= @(225, 110, 140)
        BtnDangerGradientRightPressBot= @(205, 90, 120)

        BtnSuccessGradientNormalTop    = @(110, 205, 130)
        BtnSuccessGradientNormalBot    = @(75, 180, 100)
        BtnSuccessGradientHoverTop     = @(135, 220, 150)
        BtnSuccessGradientHoverBot     = @(95, 195, 115)
        BtnSuccessGradientLeftPressTop = @(75, 175, 95)
        BtnSuccessGradientLeftPressBot = @(55, 150, 75)
        BtnSuccessGradientRightPressTop= @(95, 195, 115)
        BtnSuccessGradientRightPressBot= @(70, 170, 92)

        BtnPrimaryGradientNormalTop    = @(110, 170, 235)
        BtnPrimaryGradientNormalBot    = @(60, 120, 200)
        BtnPrimaryGradientHoverTop     = @(135, 190, 245)
        BtnPrimaryGradientHoverBot     = @(85, 145, 215)
        BtnPrimaryGradientLeftPressTop = @(55, 115, 195)
        BtnPrimaryGradientLeftPressBot = @(40, 90, 165)
        BtnPrimaryGradientRightPressTop= @(85, 145, 220)
        BtnPrimaryGradientRightPressBot= @(60, 120, 200)
    }

    # =========================================================================
    # 12. FLYFF UNIVERSE (Dark Mode)
    # =========================================================================
    FlyffUniverse = @{
        # --- Base Surfaces & Windows ---
        MainBg                      = @(26, 26, 30)      # FU near-black charcoal
        SecondaryBg                 = @(37, 37, 43)      # Elevated dark panel
        TopBarBg                    = @(20, 20, 24)      # Recessed navigation strip
        BorderColor                 = @(66, 66, 76)      # Subtle dark outline
        PanelBg                     = @(37, 37, 43)

        # --- Typography & Content ---
        PrimaryText                 = @(232, 232, 236)   # Soft white text
        HeaderTitleText             = @(245, 245, 248)
        HeaderSubtitleText          = @(150, 150, 160)   # Muted gray subtitle
        SectionHeaderText           = @(240, 185, 60)    # Signature FU gold
        AccentDanger                = @(235, 80, 80)
        StatusActive                = @(80, 200, 110)
        CurrencyPerin               = @(240, 185, 60)    # Gold perin
        CurrencyEcp                 = @(120, 200, 130)

        # --- Semantic Aliases ---
        WhiteText           = @(240, 240, 244)
        SecondaryText       = @(150, 150, 160)
        AccentBlue          = @(90, 140, 220)
        TertiaryBg          = @(46, 46, 53)
        TopBarAccent        = @(240, 185, 60)
        InputBgAlt          = @(31, 31, 36)
        SuccessColor        = @(80, 200, 110)
        SuccessColorAlt     = @(80, 200, 110)
        WarningColor        = @(240, 180, 60)
        WarningColorAlt     = @(240, 180, 60)
        ErrorColor          = @(230, 75, 75)
        ErrorColorAlt       = @(230, 75, 75)

        # --- Status & Notification Colors ---
        StatusDisconnected        = @(160, 40, 40)
        StatusDisconnectedText    = @(255, 245, 245)
        StatusFlashing            = @(255, 170, 40)
        StatusFlashingText        = @(25, 25, 30)
        StatusFlashingSelect      = @(220, 140, 30)
        NotificationProgressHigh  = @(75, 200, 110)
        NotificationProgressMid   = @(240, 185, 60)
        NotificationProgressLow   = @(230, 75, 75)
        NotificationStripWarning  = @(240, 180, 60)
        NotificationStripInfo     = @(90, 145, 225)
        NotificationStripError    = @(230, 85, 85)
        NotificationBg            = @(31, 31, 36)
        NotificationStripBg       = @(50, 50, 58)
        NotificationBtnBg         = @(55, 55, 64)
        NotificationBtnHover      = @(75, 75, 86)
        NotificationBtnPressed    = @(24, 24, 29)
        NotificationMsgText       = @(200, 200, 210)
        TooltipBorderColor        = @(150, 140, 110)

        # --- ToolTips ---
        ToolTipBg                   = @(31, 31, 36)
        ToolTipBorder               = @(200, 160, 70)
        ToolTipFg                   = @(232, 232, 236)

        # --- Flat & Standard Buttons ---
        ButtonBg                    = @(50, 50, 58)      # Dark charcoal button
        HeaderButtonBg              = @(44, 44, 51)      # Title bar gradient top stop
        BossButtonBg                = @(75, 45, 42)
        BossTimerText               = @(255, 245, 242)

        # --- TextBoxes, Inputs & Custom ComboBoxes ---
        InputBg                     = @(28, 28, 33)      # Recessed input well
        InputFg                     = @(232, 232, 236)
        ComboDropdownBg             = @(37, 37, 43)
        ComboItemText               = @(232, 232, 236)
        ComboBorderNormal           = @(72, 72, 84)
        ComboArrow                  = @(240, 185, 60)    # Gold chevron
        ComboSelectedItem           = @(160, 125, 40)    # Golden selection
        ComboSelectedText           = @(255, 255, 255)

        # --- DataGridView & Tables ---
        GridBg                      = @(26, 26, 30)
        GridBorder                  = @(52, 52, 61)
        GridSecondaryBg             = @(37, 37, 43)
        GridAlternateRow            = @(32, 32, 38)
        GridSelectedRow             = @(160, 125, 40)
        GridSelectedText            = @(255, 255, 255)
        GridHeaderBg                = @(46, 46, 53)
        GridHeaderText              = @(235, 235, 240)
        GridRowHeaderBg             = @(46, 46, 53)
        GridRowHeaderText           = @(235, 235, 240)

        # --- Tab Controls & Pages ---
        TabBg                       = @(26, 26, 30)
        TabActive                   = @(37, 37, 43)
        TabInactive                 = @(26, 26, 30)
        TabActiveText               = @(245, 245, 248)
        TabInactiveText             = @(150, 150, 160)
        TabUnderline                = @(240, 185, 60)    # Gold underline
        TabInactiveUnderline        = @(66, 66, 76)
        TabPageBg                   = @(26, 26, 30)
        TabPageText                 = @(232, 232, 236)

        # --- Switches, Trees, Lists & Progress ---
        ToggleOnColor               = @(240, 185, 60)
        ToggleOffColor              = @(140, 55, 55)
        GroupBoxBg                  = @(37, 37, 43)
        GroupBoxText                = @(235, 235, 240)
        ListItemBg                  = @(37, 37, 43)
        ListItemText                = @(232, 232, 236)
        ProgressBarBg               = @(28, 28, 33)
        ProgressBarFill             = @(240, 185, 60)    # Gold progress fill
        TreeViewBg                  = @(31, 31, 36)
        TreeViewHover               = @(55, 54, 48)
        ScrollThumbNormal           = @(52, 52, 61)
        ScrollThumbHover            = @(72, 72, 84)
        ScrollThumbActive           = @(240, 185, 60)

        # --- Button Gradients & Animated Fade States ---
        BtnGradientNormalTop        = @(54, 54, 63)
        BtnGradientNormalBot        = @(37, 37, 44)
        BtnGradientHoverTop         = @(74, 74, 86)
        BtnGradientHoverBot         = @(52, 52, 61)
        BtnGradientLeftPressTop     = @(33, 33, 39)
        BtnGradientLeftPressBot     = @(24, 24, 29)
        BtnGradientRightPressTop    = @(74, 66, 48)
        BtnGradientRightPressBot    = @(53, 47, 34)

        BtnDangerGradientNormalTop    = @(195, 70, 65)
        BtnDangerGradientNormalBot    = @(155, 45, 42)
        BtnDangerGradientHoverTop     = @(220, 90, 84)
        BtnDangerGradientHoverBot     = @(180, 58, 54)
        BtnDangerGradientLeftPressTop = @(145, 44, 40)
        BtnDangerGradientLeftPressBot = @(110, 28, 26)
        BtnDangerGradientRightPressTop= @(185, 64, 60)
        BtnDangerGradientRightPressBot= @(145, 44, 40)

        BtnSuccessGradientNormalTop    = @(75, 190, 105)
        BtnSuccessGradientNormalBot    = @(48, 150, 78)
        BtnSuccessGradientHoverTop     = @(95, 215, 125)
        BtnSuccessGradientHoverBot     = @(62, 172, 95)
        BtnSuccessGradientLeftPressTop = @(58, 145, 78)
        BtnSuccessGradientLeftPressBot = @(40, 110, 56)
        BtnSuccessGradientRightPressTop= @(72, 182, 100)
        BtnSuccessGradientRightPressBot= @(46, 145, 75)

        BtnPrimaryGradientNormalTop    = @(245, 195, 75)
        BtnPrimaryGradientNormalBot    = @(195, 145, 40)
        BtnPrimaryGradientHoverTop     = @(255, 215, 110)
        BtnPrimaryGradientHoverBot     = @(220, 170, 55)
        BtnPrimaryGradientLeftPressTop = @(190, 142, 40)
        BtnPrimaryGradientLeftPressBot = @(155, 112, 30)
        BtnPrimaryGradientRightPressTop= @(230, 180, 60)
        BtnPrimaryGradientRightPressBot= @(185, 138, 38)
    }
}

$script:Themes = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, System.Drawing.Color]]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($themeName in $rawThemes.Keys) {
    $dict = [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $source = $rawThemes[$themeName]
    foreach ($k in $source.Keys) {
        $rgb = $source[$k]
        $dict[$k] = [System.Drawing.Color]::FromArgb($rgb[0], $rgb[1], $rgb[2])
    }
    $script:Themes[$themeName] = $dict
}

function Register-CustomThemeFromConfig {
    param([switch]$Force)

    if (-not $Force -and -not $script:CustomThemeDirty -and $script:Themes.ContainsKey('Custom')) {
        return
    }

    $customDict = [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $baseTheme = $script:Themes['Dark']
    foreach ($k in $baseTheme.Keys) {
        $customDict[$k] = $baseTheme[$k]
    }

    if ($global:DashboardConfig -and $global:DashboardConfig.Config -and $global:DashboardConfig.Config['CustomTheme']) {
        $cfg = $global:DashboardConfig.Config['CustomTheme']
        foreach ($key in $cfg.Keys) {
            $valStr = [string]$cfg[$key]
            if ($valStr -match '^\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*$') {
                $r = [Math]::Min(255, [Math]::Max(0, [int]$Matches[1]))
                $g = [Math]::Min(255, [Math]::Max(0, [int]$Matches[2]))
                $b = [Math]::Min(255, [Math]::Max(0, [int]$Matches[3]))
                $customDict[$key] = [System.Drawing.Color]::FromArgb($r, $g, $b)
            }
        }
    }
    $script:Themes['Custom'] = $customDict
    $script:CustomThemeDirty = $false
}

$script:CustomThemeDirty = $true
Register-CustomThemeFromConfig -Force

if ($null -eq $global:DashboardConfig)       { $global:DashboardConfig = @{} }
if ($null -eq $global:DashboardConfig.State) { $global:DashboardConfig.State = @{} }

$script:ActiveThemeName = [string]$global:DashboardConfig.State.CurrentTheme
if (-not $script:Themes.ContainsKey($script:ActiveThemeName)) {
    $script:ActiveThemeName = 'OriginalGray'
    $global:DashboardConfig.State.CurrentTheme = 'OriginalGray'
}
$script:ActiveCache = $script:Themes[$script:ActiveThemeName]

function Save-CustomThemeColors {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]$ColorMap,

        [switch]$NoConfigWrite
    )

    if (-not $NoConfigWrite) {
        if (-not $global:DashboardConfig.Config) { $global:DashboardConfig.Config = [ordered]@{} }
        if (-not $global:DashboardConfig.Config['CustomTheme']) { $global:DashboardConfig.Config['CustomTheme'] = [ordered]@{} }
    }

    $customDict = [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]::new([System.StringComparer]::OrdinalIgnoreCase)
    
    foreach ($key in $ColorMap.Keys) {
        $color = $ColorMap[$key]
        $customDict[$key] = $color
        if (-not $NoConfigWrite) {
            $global:DashboardConfig.Config['CustomTheme'][$key] = "$($color.R),$($color.G),$($color.B)"
        }
    }

    $script:Themes['Custom'] = $customDict
    $script:CustomThemeDirty = $false

    if (-not $NoConfigWrite -and (Test-DashboardCommand 'WriteConfig')) {
        WriteConfig
    }
}


function Get-Theme {
    [CmdletBinding(DefaultParameterSetName = 'Color')]
    param(
        [Parameter(ParameterSetName = 'Color', Position = 0)]
        [string]$Color,

        [Parameter(ParameterSetName = 'Font')]
        [string]$Font = 'DefaultSize',

        [Parameter(ParameterSetName = 'Font')]
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular,

        [Parameter(ParameterSetName = 'Font')]
        [string]$Family = 'Segoe UI'
    )

    if ($PSCmdlet.ParameterSetName -eq 'Font') {
        $sz = if ($script:FontSizes -and $script:FontSizes.ContainsKey($Font)) { $script:FontSizes[$Font] } else { 9.0 }
        return [System.Drawing.Font]::new($Family, $sz, $Style)
    }

    if ($Color) {
        if ($Color -eq 'Transparent') { return [System.Drawing.Color]::Transparent }
        if ($Color -eq 'Empty') { return [System.Drawing.Color]::Empty }
        $val = [System.Drawing.Color]::Empty
        if ($script:ActiveCache.TryGetValue($Color, [ref]$val)) { return $val }
        return [System.Drawing.Color]::Gray
    }

    return $script:ActiveCache
}

function New-ColorGradient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Rectangle]$Rectangle,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$StartColor,

        [Parameter(Mandatory = $true)]
        [System.Drawing.Color]$EndColor,

        [System.Drawing.Drawing2D.LinearGradientMode]$Mode = [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
    )

    return [System.Drawing.Drawing2D.LinearGradientBrush]::new($Rectangle, $StartColor, $EndColor, $Mode)
}

function New-ThemedToolTip {
    [CmdletBinding()]
    param()

    $toolTip = [System.Windows.Forms.ToolTip]@{
        AutoPopDelay = 5000
        InitialDelay = 100
        ReshowDelay  = 10
        ShowAlways   = $true
        OwnerDraw    = $true
    }
    $toolTip | Add-Member -MemberType NoteProperty -Name 'TipFont' -Value (Get-UIFont 'Segoe UI' 9) -Force

    $toolTip.Add_Draw({
        param($sender, $e)
        $c = Get-Theme
        $bgBrush   = [System.Drawing.SolidBrush]::new($c['ToolTipBg'])
        $borderPen = [System.Drawing.Pen]::new($c['ToolTipBorder'])
        $textBrush = [System.Drawing.SolidBrush]::new($c['ToolTipFg'])

        $e.Graphics.FillRectangle($bgBrush, $e.Bounds)
        $e.Graphics.DrawRectangle($borderPen, $e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width - 1, $e.Bounds.Height - 1)
        $e.Graphics.DrawString($e.ToolTipText, $sender.TipFont, $textBrush, 3, 3, [System.Drawing.StringFormat]::GenericTypographic)

        $bgBrush.Dispose()
        $borderPen.Dispose()
        $textBrush.Dispose()
    })

    $toolTip.Add_Popup({
        param($sender, $e)
        $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
        $s = $g.MeasureString($sender.GetToolTip($e.AssociatedControl), $sender.TipFont, [System.Drawing.PointF]::new(0, 0), [System.Drawing.StringFormat]::GenericTypographic)
        $g.Dispose()
        $e.ToolTipSize = [System.Drawing.Size]::new([int]$s.Width + 12, [int]$s.Height + 8)
    })

    return $toolTip
}

function Register-ThemedTopBarPaint([System.Windows.Forms.Control]$Control) {
    if ($null -eq $Control) { return }
    $Control.Add_Paint({
        param($src, $e)
        $c = Get-Theme
        $rect = [System.Drawing.Rectangle]::new(0, 0, $src.Width, $src.Height)
        $c1   = $c['HeaderButtonBg']
        $c2   = $c['TopBarBg']
        $grad = New-ColorGradient -Rectangle $rect -StartColor $c1 -EndColor $c2 -Mode Vertical
        $e.Graphics.FillRectangle($grad, $rect)
        $grad.Dispose()

        $pen = [System.Drawing.Pen]::new($c['BorderColor'])
        $e.Graphics.DrawLine($pen, 0, $src.Height - 1, $src.Width, $src.Height - 1)
        $pen.Dispose()
    })
}

function Register-ThemedResizeGripPaint([System.Windows.Forms.Control]$Control) {
    if ($null -eq $Control) { return }
    $Control.Add_Paint({
        param($src, $e)
        $c = Get-Theme
        $pen = [System.Drawing.Pen]::new($c['BorderColor'])
        $e.Graphics.DrawLine($pen, 0, 0, $src.Width, 0)
        $pen.Dispose()
    })
}

function Register-ThemedFlatButtonPaint([System.Windows.Forms.Button]$Button) {
    if ($null -eq $Button) { return }
    $Button.Add_Paint({
        param($src, $e)
        if ($src.FlatStyle -eq [System.Windows.Forms.FlatStyle]::Flat) {
            $c = Get-Theme
            $bgBrush = [System.Drawing.SolidBrush]::new($src.BackColor)
            $e.Graphics.FillRectangle($bgBrush, 0, 0, $src.Width, $src.Height)

            $tagStr = if ($src.Tag -is [System.Collections.IDictionary]) { [string]$src.Tag['Type'] } else { [string]$src.Tag }
            $fg = if ($tagStr -eq 'Boss') { $c['BossTimerText'] } else { $src.ForeColor }

            $textBrush = [System.Drawing.SolidBrush]::new($fg)
            $textFormat = [System.Drawing.StringFormat]::new()
            $textFormat.Alignment = [System.Drawing.StringAlignment]::Center
            $textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
            $img = $src.Image
            if ($null -ne $img) {
                $imgH = [int][Math]::Min([int]$img.Height, [Math]::Max(1, ($src.Height - 4)))
                $imgW = [int][Math]::Round([double]$img.Width * ($imgH / [Math]::Max(1.0, [double]$img.Height)))
                if ($imgW -lt 1) { $imgW = 1 }
                if ($imgH -lt 1) { $imgH = 1 }
                $imgX = [int](($src.Width - $imgW) / 2)
                $imgY = [int](($src.Height - $imgH) / 2)
                if (-not [string]::IsNullOrEmpty($src.Text)) {
                    $imgX = 7
                    $e.Graphics.DrawImage($img, $imgX, $imgY, $imgW, $imgH)
                    $tx = [float]($imgX + $imgW + 2)
                    $tw = [float]([Math]::Max(4, ($src.Width - $imgX - $imgW - 4)))
                    $e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new($tx, [float]0, $tw, [float]$src.Height), $textFormat)
                } else {
                    $e.Graphics.DrawImage($img, $imgX, $imgY, $imgW, $imgH)
                }
            } else {
                $e.Graphics.DrawString($src.Text, $src.Font, $textBrush, [System.Drawing.RectangleF]::new(0, 0, $src.Width, $src.Height), $textFormat)
            }

            $borderPen = [System.Drawing.Pen]::new($c['BorderColor'])
            $e.Graphics.DrawRectangle($borderPen, 0, 0, $src.Width - 1, $src.Height - 1)

            $bgBrush.Dispose()
            $textBrush.Dispose()
            $borderPen.Dispose()
            $textFormat.Dispose()
        }
    })
}

function Register-ThemedCheckBoxPaint([System.Windows.Forms.CheckBox]$CheckBox) {
    if ($null -eq $CheckBox) { return }
    if ($CheckBox.PSObject.Properties['ThemedCheckBoxPaint']) { return }
    $CheckBox | Add-Member -MemberType NoteProperty -Name 'ThemedCheckBoxPaint' -Value $true -Force

    $CheckBox.Add_Paint({
        param($src, $e)
        $c = Get-Theme
        $g = $e.Graphics

        $bg = $src.BackColor
        if ($bg.A -lt 255) {
            if ($null -ne $src.Parent -and $src.Parent.BackColor.A -eq 255) {
                $bg = $src.Parent.BackColor
            }
        }

        $bgBrush = [System.Drawing.SolidBrush]::new($bg)
        $g.FillRectangle($bgBrush, 0, 0, $src.Width, $src.Height)
        $bgBrush.Dispose()

        $boxSize = 14
        $boxX = 2
        $boxY = [Math]::Max(0, [int](($src.Height - $boxSize) / 2))

        $fillColor = if ($src.Checked) { $c['AccentBlue'] } else { $c['GridSecondaryBg'] }
        $fillBrush = [System.Drawing.SolidBrush]::new($fillColor)
        $g.FillRectangle($fillBrush, $boxX, $boxY, $boxSize, $boxSize)
        $fillBrush.Dispose()

        $borderPen = [System.Drawing.Pen]::new($c['BorderColor'])
        $g.DrawRectangle($borderPen, $boxX, $boxY, $boxSize, $boxSize)
        $borderPen.Dispose()

        if ($src.CheckState -eq [System.Windows.Forms.CheckState]::Indeterminate) {
            $barBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
            $g.FillRectangle($barBrush, $boxX + 3, $boxY + 6, 8, 2)
            $barBrush.Dispose()
        } elseif ($src.Checked) {
            $checkPen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, 2)
            $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
            $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $g.DrawLine($checkPen, $boxX + 3, $boxY + 7, $boxX + 6, $boxY + 10)
            $g.DrawLine($checkPen, $boxX + 6, $boxY + 10, $boxX + 11, $boxY + 4)
            $checkPen.Dispose()
        }

        $textColor = $src.ForeColor
        if (-not $src.Enabled) {
            $textColor = [System.Drawing.Color]::FromArgb(
                [int](($textColor.R + $bg.R) / 2),
                [int](($textColor.G + $bg.G) / 2),
                [int](($textColor.B + $bg.B) / 2)
            )
        }

        $flags = [System.Windows.Forms.TextFormatFlags]::Left -bor
                 [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::SingleLine -bor
                 [System.Windows.Forms.TextFormatFlags]::NoPrefix
        $tx = [Math]::Max($src.Padding.Left, 18)
        $tr = [System.Drawing.Rectangle]::new($tx, 0, [Math]::Max(1, $src.Width - $tx), $src.Height)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $src.Text, $src.Font, $tr, $textColor, $flags)
    })
}

function Register-ThemedGroupBoxPaint([System.Windows.Forms.GroupBox]$GroupBox) {
    if ($null -eq $GroupBox) { return }
    if ($GroupBox.PSObject.Properties['ThemedGroupBoxPaint']) { return }
    $GroupBox | Add-Member -MemberType NoteProperty -Name 'ThemedGroupBoxPaint' -Value $true -Force

    $GroupBox.Add_Paint({
        param($src, $e)
        $c = Get-Theme
        $g = $e.Graphics

        $bg = $src.BackColor
        if ($bg.A -lt 255) {
            if ($null -ne $src.Parent -and $src.Parent.BackColor.A -eq 255) {
                $bg = $src.Parent.BackColor
            } else {
                $bg = $c['GroupBoxBg']
            }
        }

        $bgBrush = [System.Drawing.SolidBrush]::new($bg)
        $g.FillRectangle($bgBrush, 0, 0, $src.Width, $src.Height)
        $bgBrush.Dispose()

        $borderPen = [System.Drawing.Pen]::new($c['BorderColor'])
        $g.DrawRectangle($borderPen, 0, 0, $src.Width - 1, $src.Height - 1)
        $borderPen.Dispose()

        if (-not [string]::IsNullOrEmpty($src.Text)) {
            $measFlags = [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
            $sz = [System.Windows.Forms.TextRenderer]::MeasureText($src.Text, $src.Font, [System.Drawing.Size]::new(0, 0), $measFlags)

            $gapBrush = [System.Drawing.SolidBrush]::new($bg)
            $g.FillRectangle($gapBrush, 8, 0, $sz.Width + 6, [Math]::Max($sz.Height, $src.Font.Height))
            $gapBrush.Dispose()

            $fg = $src.ForeColor
            if (-not $src.Enabled) {
                $fg = [System.Drawing.Color]::FromArgb(
                    [int](($fg.R + $bg.R) / 2),
                    [int](($fg.G + $bg.G) / 2),
                    [int](($fg.B + $bg.B) / 2)
                )
            }

            $flags = [System.Windows.Forms.TextFormatFlags]::Left -bor
                     [System.Windows.Forms.TextFormatFlags]::Top -bor
                     [System.Windows.Forms.TextFormatFlags]::SingleLine -bor
                     [System.Windows.Forms.TextFormatFlags]::NoPrefix
            $tr = [System.Drawing.Rectangle]::new(11, 0, $src.Width - 12, [Math]::Max($sz.Height, $src.Font.Height))
            [System.Windows.Forms.TextRenderer]::DrawText($g, $src.Text, $src.Font, $tr, $fg, $flags)
        }
    })
}

function Get-ThemedFitFont {
    [CmdletBinding()]
    param(
        [string]$Text,
        [int]$MaxWidth,
        [System.Drawing.Font]$Base,
        [double]$Min = 5.0
    )
    if ($null -eq $Base) { return $Base }
    if ([string]::IsNullOrEmpty($Text) -or $MaxWidth -lt 8) { return $Base }
    try {
        if ([System.Windows.Forms.TextRenderer]::MeasureText($Text, $Base).Width -le $MaxWidth) { return $Base }
    } catch { return $Base }

    if ($null -eq $script:ThemedFitFonts) { $script:ThemedFitFonts = @{} }
    $size = [double]$Base.SizeInPoints
    $fit = $Base
    while (($size - 0.5) -ge ($Min - 0.001)) {
        $size = [Math]::Round($size - 0.5, 1)
        $key = "$($Base.Name)|$($Base.Style)|$($size)"
        if (-not $script:ThemedFitFonts.ContainsKey($key)) {
            $script:ThemedFitFonts[$key] = New-Object System.Drawing.Font($Base.Name, $size, $Base.Style)
        }
        $f = $script:ThemedFitFonts[$key]
        $fit = $f
        $ok = $false
        try { $ok = ([System.Windows.Forms.TextRenderer]::MeasureText($Text, $f).Width -le $MaxWidth) } catch { $ok = $false }
        if ($ok) { return $f }
    }
    return $fit
}

function Set-ThemedFitText {
    [CmdletBinding()]
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Text = '',
        [double]$Min = 5.0
    )
    if ($null -eq $Control) { return }
    try { $Control.Text = $Text } catch { return }
    try {
        $prop = $Control.PSObject.Properties['ThemedFitBaseFont']
        if (-not $prop -or $null -eq $prop.Value) {
            $Control | Add-Member -MemberType NoteProperty -Name 'ThemedFitBaseFont' -Value $Control.Font -Force
            $prop = $Control.PSObject.Properties['ThemedFitBaseFont']
        }
        $base = $prop.Value
        if ($null -eq $base) { return }
        $w = [int]$Control.ClientSize.Width - 4
        $fit = Get-ThemedFitFont -Text $Text -MaxWidth $w -Base $base -Min $Min
        if ($null -ne $fit) { $Control.Font = $fit }
    } catch { }
}

function Get-ThemedGridFitBaseFont {
    [CmdletBinding()]
    param(
        $Grid,
        [int]$ColumnIndex = -1
    )
    if ($null -eq $Grid) { return $null }
    $f = $null
    if ($ColumnIndex -ge 0) {
        try {
            $col = $Grid.Columns[$ColumnIndex]
            if ($null -ne $col) { $f = $col.DefaultCellStyle.Font }
        } catch { $f = $null }
        if ($null -ne $f) { return $f }
    }
    try { $f = $Grid.DefaultCellStyle.Font } catch { $f = $null }
    if ($null -ne $f) { return $f }
    try { $f = $Grid.Font } catch { $f = $null }
    return $f
}

function Register-ThemedGridFitText([System.Windows.Forms.DataGridView]$Grid) {
    if ($null -eq $Grid) { return }
    if ($Grid.PSObject.Properties['ThemedGridFitText']) { return }
    $Grid | Add-Member -MemberType NoteProperty -Name 'ThemedGridFitText' -Value $true -Force

    $Grid.Add_CellFormatting({
        param($src, $e)
        try {
            if ($null -eq $e -or $null -eq $e.Value) { return }
            if ($e.Value -isnot [string]) { return }
            $txt = [string]$e.Value
            if ($txt -eq '') { return }
            $base = $null
            try { $base = Get-ThemedGridFitBaseFont -Grid $src -ColumnIndex ([int]$e.ColumnIndex) } catch { $base = $null }
            if ($null -eq $base) { return }
            $w = [int]$e.CellBounds.Width - 6
            if ($w -lt 8) { return }
            $fit = Get-ThemedFitFont -Text $txt -MaxWidth $w -Base $base
            if ($null -ne $fit) { $e.CellStyle.Font = $fit }
        } catch { }
    })
}

function Register-ThemedGridToolTip([System.Windows.Forms.DataGridView]$Grid) {
    if ($null -eq $Grid) { return }
    if ($Grid.PSObject.Properties['ThemedGridToolTip']) { return }
    $Grid | Add-Member -MemberType NoteProperty -Name 'ThemedGridToolTip' -Value $true -Force

    $tip = New-ThemedToolTip
    $Grid | Add-Member -MemberType NoteProperty -Name 'ThemedGridTip' -Value $tip -Force
    $Grid | Add-Member -MemberType NoteProperty -Name 'ThemedGridTipText' -Value $null -Force
    $Grid.ShowCellToolTips = $false

    $Grid.Add_MouseMove({
        param($src, $e)
        try {
            $hit = $src.HitTest($e.X, $e.Y)
            $text = $null

            if ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::ColumnHeader) {
                $noHdr = $src.PSObject.Properties['ThemedGridTipNoHeaders']
                if ($noHdr -and $noHdr.Value) { $text = $null }
                elseif ($hit.ColumnIndex -ge 0) { $text = [string]$src.Columns[$hit.ColumnIndex].ToolTipText }
            } elseif ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::RowHeader) {
                if ($hit.RowIndex -ge 0) { $text = [string]$src.Rows[$hit.RowIndex].ToolTipText }
            } elseif ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::Cell) {
                $noCells = $src.PSObject.Properties['ThemedGridTipNoCells']
                if (-not ($noCells -and $noCells.Value)) {
                    if ($hit.RowIndex -ge 0 -and $hit.ColumnIndex -ge 0) {
                        $cell = $src.Rows[$hit.RowIndex].Cells[$hit.ColumnIndex]
                        $text = [string]$cell.ToolTipText
                        if ([string]::IsNullOrEmpty($text) -and $null -ne $cell.OwningColumn) {
                            $text = [string]$cell.OwningColumn.ToolTipText
                        }
                        if ([string]::IsNullOrEmpty($text) -and $null -ne $cell.OwningRow) {
                            $text = [string]$cell.OwningRow.ToolTipText
                        }
                    }
                }
            }

            $curProp = $src.PSObject.Properties['ThemedGridTipText']
            $cur = if ($curProp) { $curProp.Value } else { $null }

            if ([string]::IsNullOrEmpty($text)) {
                if (-not [string]::IsNullOrEmpty($cur)) {
                    $tipHide = $src.PSObject.Properties['ThemedGridTip']
                    if ($tipHide -and $tipHide.Value) { try { $tipHide.Value.Hide([System.Windows.Forms.Control]$src) } catch { } }
                    $src | Add-Member -MemberType NoteProperty -Name 'ThemedGridTipText' -Value $null -Force
                }
                return
            }

            if ($text -eq $cur) { return }

            $tipProp = $src.PSObject.Properties['ThemedGridTip']
            if ($tipProp -and $tipProp.Value) {
                $src | Add-Member -MemberType NoteProperty -Name 'ThemedGridTipText' -Value $text -Force
                $tipProp.Value.SetToolTip([System.Windows.Forms.Control]$src, $text)
                $pos = [System.Drawing.Point]::new($e.X + 2, $e.Y + 18)
                $tipProp.Value.Show($text, [System.Windows.Forms.Control]$src, $pos)
            }
        } catch { }
    })

    $Grid.Add_MouseLeave({
        param($src, $e)
        $tipProp = $src.PSObject.Properties['ThemedGridTip']
        if ($tipProp -and $tipProp.Value) {
            try { $tipProp.Value.Hide([System.Windows.Forms.Control]$src) } catch { }
        }
        $src | Add-Member -MemberType NoteProperty -Name 'ThemedGridTipText' -Value $null -Force
    })

    $Grid.Add_Scroll({
        param($src, $e)
        $tipProp = $src.PSObject.Properties['ThemedGridTip']
        if ($tipProp -and $tipProp.Value) {
            try { $tipProp.Value.Hide([System.Windows.Forms.Control]$src) } catch { }
        }
        $src | Add-Member -MemberType NoteProperty -Name 'ThemedGridTipText' -Value $null -Force
    })

    $Grid.Add_Disposed({
        param($src, $e)
        $tipProp = $src.PSObject.Properties['ThemedGridTip']
        if ($tipProp -and $tipProp.Value) {
            try { $tipProp.Value.Dispose() } catch { }
        }
    })
}

function Register-ThemedButtonFadeEffect {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Button]$Button,

        [ValidateSet('Normal', 'Danger', 'Success', 'Primary')]
        [string]$Variant = 'Normal'
    )

    if ($Variant -ne 'Normal') { $Button.Tag = $Variant }
    if (-not (Test-DashboardCommand 'Register-ButtonFadeEffect')) { return }

    $c = Get-Theme
    $prefix = switch ($Variant) {
        'Danger'  { 'BtnDanger' }
        'Success' { 'BtnSuccess' }
        'Primary' { 'BtnPrimary' }
        default   { 'Btn' }
    }

    Register-ButtonFadeEffect -Button $Button `
        -NormalStartColor       $c["${prefix}GradientNormalTop"] `
        -NormalEndColor         $c["${prefix}GradientNormalBot"] `
        -HoverStartColor        $c["${prefix}GradientHoverTop"] `
        -HoverEndColor          $c["${prefix}GradientHoverBot"] `
        -LeftPressedStartColor  $c["${prefix}GradientLeftPressTop"] `
        -LeftPressedEndColor    $c["${prefix}GradientLeftPressBot"] `
        -RightPressedStartColor $c["${prefix}GradientRightPressTop"] `
        -RightPressedEndColor   $c["${prefix}GradientRightPressBot"]
}

function ApplyControlInternal([System.Windows.Forms.Control]$Control, [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]$c) {
    if ($null -eq $Control) { return }

    $fullName = $Control.GetType().FullName

    if ($fullName.StartsWith('Custom.', [System.StringComparison]::Ordinal)) {
        switch ($fullName) {
            'Custom.DarkComboBox' {
                $Control.BackgroundColor       = $c['ComboDropdownBg']
                $Control.TextColor             = $c['ComboItemText']
                $Control.BorderColor           = $c['ComboBorderNormal']
                $Control.ArrowColor            = $c['ComboArrow']
                $Control.SelectedItemBackColor = $c['ComboSelectedItem']
                $Control.SelectedItemForeColor = $c['ComboSelectedText']
                $Control.DropDownBackColor     = $c['ComboDropdownBg']
                $Control.BackColor             = $c['ComboDropdownBg']
                $Control.ForeColor             = $c['ComboItemText']
                $Control.Invalidate()
                return
            }
            'Custom.DarkTabControl' {
                $Control.TabBgColor                = $c['TabBg']
                $Control.TabActiveBgColor          = $c['TabActive']
                $Control.TabInactiveBgColor        = $c['TabInactive']
                $Control.TabActiveTextColor        = $c['TabActiveText']
                $Control.TabInactiveTextColor      = $c['TabInactiveText']
                $Control.TabUnderlineColor         = $c['TabUnderline']
                $Control.TabInactiveUnderlineColor = $c['TabInactiveUnderline']
                $Control.BackColor                 = $c['TabBg']
                $Control.ForeColor                 = $c['TabActiveText']
                $Control.Invalidate()
                return
            }
            'Custom.Toggle' {
                $Control.OnColor   = $c['ToggleOnColor']
                $Control.OffColor  = $c['ToggleOffColor']
                $Control.BackColor = [System.Drawing.Color]::Transparent
                $Control.Invalidate()
                return
            }
            'Custom.TextProgressBar' {
                $Control.FillColor = $c['ProgressBarFill']
                $Control.BackColor = $c['ProgressBarBg']
                $Control.Invalidate()
                return
            }
            'Custom.DarkTreeView' {
                $Control.BackColor      = $c['TreeViewBg']
                $Control.ForeColor      = $c['PrimaryText']
                $Control.SelectionColor = $c['TreeViewHover']
                $Control.Invalidate()
                return
            }
            'Custom.DarkScrollBar' {
                $Control.BackColor        = $c['MainBg']
                $Control.ThumbNormalColor = $c['ScrollThumbNormal']
                $Control.ThumbHoverColor  = $c['ScrollThumbHover']
                $Control.ThumbActiveColor = $c['ScrollThumbActive']
                $Control.Invalidate()
                return
            }
        }
    }

    if ($Control -is [System.Windows.Forms.Label]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        switch ([string]$Control.Tag) {
            'Title'    { $Control.ForeColor = $c['HeaderTitleText'] }
            'Subtitle' { $Control.ForeColor = $c['HeaderSubtitleText'] }
            'Muted'    { $Control.ForeColor = $c['HeaderSubtitleText'] }
            'Section'  { $Control.ForeColor = $c['SectionHeaderText'] }
            'Danger'   { $Control.ForeColor = $c['AccentDanger'] }
            'Accent'   { $Control.ForeColor = $c['AccentDanger'] }
            'Success'  { $Control.ForeColor = $c['StatusActive'] }
            'Status'   { $Control.ForeColor = $c['StatusActive'] }
            'Perin'    { $Control.ForeColor = $c['CurrencyPerin'] }
            'Ecp'      { $Control.ForeColor = $c['CurrencyEcp'] }
            default    { $Control.ForeColor = $c['PrimaryText'] }
        }
        return
    }

    if ($Control -is [System.Windows.Forms.Button]) {
        $tagStr = if ($Control.Tag -is [System.Collections.IDictionary]) { [string]$Control.Tag['Type'] } else { [string]$Control.Tag }
        $Control.ForeColor = if ($tagStr -eq 'Boss') { $c['BossTimerText'] } else { $c['PrimaryText'] }
        $Control.FlatAppearance.BorderColor = $c['BorderColor']

        $fadeProp = $Control.PSObject.Properties['FadeState']
        if ($fadeProp -and $fadeProp.Value) {
            $fs = $fadeProp.Value
            $prefix = switch ($tagStr) {
                'Danger'  { 'BtnDanger' }
                'Success' { 'BtnSuccess' }
                'Primary' { 'BtnPrimary' }
                default   { 'Btn' }
            }

            $top = $c["${prefix}GradientNormalTop"]
            $bot = $c["${prefix}GradientNormalBot"]

            $fs.NormalStartColor       = $top
            $fs.NormalEndColor         = $bot
            $fs.HoverStartColor        = $c["${prefix}GradientHoverTop"]
            $fs.HoverEndColor          = $c["${prefix}GradientHoverBot"]
            $fs.LeftPressedStartColor  = $c["${prefix}GradientLeftPressTop"]
            $fs.LeftPressedEndColor    = $c["${prefix}GradientLeftPressBot"]
            $fs.RightPressedStartColor = $c["${prefix}GradientRightPressTop"]
            $fs.RightPressedEndColor   = $c["${prefix}GradientRightPressBot"]

            $fs.TargetStartColor = $top
            $fs.TargetEndColor   = $bot
            $fs.CurrentStartR    = [double]$top.R
            $fs.CurrentStartG    = [double]$top.G
            $fs.CurrentStartB    = [double]$top.B
            $fs.CurrentEndR      = [double]$bot.R
            $fs.CurrentEndG      = [double]$bot.G
            $fs.CurrentEndB      = [double]$bot.B

            $Control.Invalidate()
        } else {
            $Control.BackColor = switch ($tagStr) {
                'Danger'       { $c['BtnDangerGradientNormalTop'] }
                'Success'      { $c['BtnSuccessGradientNormalTop'] }
                'Primary'      { $c['BtnPrimaryGradientNormalTop'] }
                'Boss'         { $c['BossButtonBg'] }
                'HeaderButton' { $c['HeaderButtonBg'] }
                default        { $c['ButtonBg'] }
            }
        }
        return
    }

    if ($Control -is [System.Windows.Forms.Panel]) {
        $Control.BackColor = switch ([string]$Control.Tag) {
            'MainBg' { $c['MainBg'] }
            'TopBar' { $c['TopBarBg'] }
            default  { $c['SecondaryBg'] }
        }
        $Control.Invalidate()
        return
    }

    if ($Control -is [System.Windows.Forms.TableLayoutPanel]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        return
    }

    if ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.ComboBox]) {
        $Control.BackColor = $c['InputBg']
        $Control.ForeColor = $c['InputFg']
        return
    }

    if ($Control -is [System.Windows.Forms.DataGridView]) {
        $Control.EnableHeadersVisualStyles = $false
        $Control.BackgroundColor           = $c['GridBg']
        $Control.GridColor                 = $c['GridBorder']

        $d = $Control.DefaultCellStyle
        $d.BackColor          = $c['GridSecondaryBg']
        $d.ForeColor          = $c['PrimaryText']
        $d.SelectionBackColor = $c['GridSelectedRow']
        $d.SelectionForeColor = $c['GridSelectedText']

        $a = $Control.AlternatingRowsDefaultCellStyle
        $a.BackColor          = $c['GridAlternateRow']
        $a.ForeColor          = $c['PrimaryText']
        $a.SelectionBackColor = $c['GridSelectedRow']
        $a.SelectionForeColor = $c['GridSelectedText']

        $ch = $Control.ColumnHeadersDefaultCellStyle
        $ch.BackColor          = $c['GridHeaderBg']
        $ch.ForeColor          = $c['GridHeaderText']
        $ch.SelectionBackColor = $c['GridSelectedRow']

        $rh = $Control.RowHeadersDefaultCellStyle
        $rh.BackColor          = $c['GridRowHeaderBg']
        $rh.ForeColor          = $c['GridRowHeaderText']

        foreach ($col in $Control.Columns) {
            if ($col -is [System.Windows.Forms.DataGridViewButtonColumn]) {
                $col.DefaultCellStyle.BackColor          = $c['ButtonBg']
                $col.DefaultCellStyle.ForeColor          = $c['PrimaryText']
                $col.DefaultCellStyle.SelectionBackColor = $c['ComboSelectedItem']
                $col.DefaultCellStyle.SelectionForeColor = $c['ComboSelectedText']
            }
        }

        if ($Control.PSObject.Properties['CheckBoxCheckedColor']) {
            $Control.CheckBoxCheckedColor   = $c['AccentBlue']
            $Control.CheckBoxUncheckedColor = $c['GridSecondaryBg']
            $Control.CheckBoxBorderColor    = $c['BorderColor']
            $Control.CheckBoxCheckColor     = [System.Drawing.Color]::White
        }
        Register-ThemedGridToolTip $Control
        Register-ThemedGridFitText $Control
        $Control.Invalidate()
        return
    }

    if ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $c['TabPageBg']
        $Control.ForeColor = $c['TabPageText']
        return
    }

    if ($Control -is [System.Windows.Forms.TabControl]) {
        $Control.BackColor = $c['TabBg']
        $Control.ForeColor = $c['TabActiveText']
        return
    }

    if ($Control -is [System.Windows.Forms.CheckBox]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        $Control.ForeColor = if ($Control.Tag -eq 'Muted' -or $Control.Tag -eq 'Subtitle') { $c['HeaderSubtitleText'] } else { $c['PrimaryText'] }
        $Control.FlatAppearance.BorderColor       = $c['BorderColor']
        $Control.FlatAppearance.CheckedBackColor   = $c['ComboSelectedItem']
        $Control.FlatAppearance.MouseOverBackColor = $c['ButtonBg']
        Register-ThemedCheckBoxPaint $Control
        return
    }

    if ($Control -is [System.Windows.Forms.GroupBox]) {
        $Control.BackColor = $c['GroupBoxBg']
        $Control.ForeColor = $c['GroupBoxText']
        Register-ThemedGroupBoxPaint $Control
        return
    }

    if ($Control -is [System.Windows.Forms.ListBox]) {
        $Control.BackColor = $c['ListItemBg']
        $Control.ForeColor = $c['ListItemText']
        return
    }

    if ($Control -is [System.Windows.Forms.TreeView]) {
        $Control.BackColor = $c['TreeViewBg']
        $Control.ForeColor = $c['PrimaryText']
        return
    }

    if ($Control -is [System.Windows.Forms.ProgressBar]) {
        $Control.BackColor = $c['ProgressBarBg']
        $Control.ForeColor = $c['ProgressBarFill']
        return
    }

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $c['MainBg']
        $Control.ForeColor = $c['PrimaryText']
        return
    }

    if ($Control -is [System.Windows.Forms.ToolStrip]) {
        $darkRendererType = 'Custom.DarkRenderer' -as [Type]
        if ($script:ActiveThemeName -eq 'OriginalGray' -and $null -ne $darkRendererType) {
            $Control.Renderer = [Activator]::CreateInstance($darkRendererType)
        } else {
            $Control.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
        }

        $Control.BackColor = $c['SecondaryBg']
        $Control.ForeColor = $c['PrimaryText']

        $applyItems = {
            param($items)
            foreach ($item in $items) {
                if ($item -is [System.Windows.Forms.ToolStripItem]) {
                    $item.BackColor = $c['SecondaryBg']
                    $item.ForeColor = $c['PrimaryText']
                    if ($item -is [System.Windows.Forms.ToolStripDropDownItem] -and $item.HasDropDownItems) {
                        & $applyItems $item.DropDownItems
                        try {
                            $dd = $item.DropDown
                            if ($null -ne $dd -and -not $dd.IsDisposed) {
                                ApplyControlInternal $dd $c
                            }
                        } catch { }
                    }
                }
            }
        }
        & $applyItems $Control.Items
        return
    }
}

function ApplyControlTheme($RootControl, [System.Collections.Generic.Dictionary[string, System.Drawing.Color]]$Colors) {
    if ($null -eq $RootControl -or ($RootControl -isnot [System.Windows.Forms.Control])) { return }

    $queue = [System.Collections.Generic.Queue[System.Windows.Forms.Control]]::new()
    $queue.Enqueue($RootControl)

    while ($queue.Count -gt 0) {
        $ctrl = $queue.Dequeue()
        ApplyControlInternal $ctrl $Colors

        if ($ctrl.ContextMenuStrip) {
            ApplyControlInternal $ctrl.ContextMenuStrip $Colors
        }

        $childCollection = $ctrl.Controls
        if ($childCollection -and $childCollection.Count -gt 0) {
            for ($i = 0; $i -lt $childCollection.Count; $i++) {
                $child = $childCollection[$i]
                if ($child -is [System.Windows.Forms.Control]) {
                    $queue.Enqueue($child)
                }
            }
        }
    }

    if ($RootControl -is [System.Windows.Forms.Control]) {
        $RootControl.Invalidate($true)
    }
}

if (-not $global:DashboardConfig.Resources.OpenThemeableForms) {
    $global:DashboardConfig.Resources.OpenThemeableForms = [System.Collections.Generic.List[System.Windows.Forms.Form]]::new()
}

function Register-OpenForm([System.Windows.Forms.Form]$Form) {
    if ($null -eq $Form -or $Form.IsDisposed) { return }
    if (-not $global:DashboardConfig.Resources.OpenThemeableForms.Contains($Form)) {
        $global:DashboardConfig.Resources.OpenThemeableForms.Add($Form)
        $Form.Add_FormClosed({
            param($src, $e)
            if ($global:DashboardConfig.Resources.OpenThemeableForms) {
                $global:DashboardConfig.Resources.OpenThemeableForms.Remove($src)
            }
        })
    }
}

function Set-Theme {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [object]$ThemeOrTarget,

        [Parameter()]
        [string]$Name,

        [Parameter()]
        [System.Windows.Forms.Control]$Target,

        [Parameter()]
        [switch]$Next,

        [Parameter()]
        [switch]$List,

        [Parameter()]
        [switch]$NoSave
    )

    $lightNames = @('ArcticFrost','DeepOceanLight','FlyffClassic','Light','QuartzStone','SunnyMeadow','TerraCotta')
    $darkNames  = @('CrimsonTwilight','CyberViolet','Dark','DeepOcean','FlyffUniverse','ForestSage','MidnightSlate','OriginalGray')
    $light = @($lightNames | Where-Object { $script:Themes.ContainsKey($_) } | Sort-Object)
    $dark  = @($darkNames  | Where-Object { $script:Themes.ContainsKey($_) } | Sort-Object)
    $custom = @('Custom' | Where-Object { $script:Themes.ContainsKey($_) })
    $keys = [string[]]@($light + $dark + $custom)

    if ($List) { return $keys }

    if ($Next) {
        $idx = [array]::IndexOf($keys, $script:ActiveThemeName)
        $Name = $keys[($idx + 1) % $keys.Length]
    }

    if ($null -ne $ThemeOrTarget) {
        if ($ThemeOrTarget -is [System.Windows.Forms.Control]) {
            $Target = $ThemeOrTarget
        } elseif ($ThemeOrTarget -is [string]) {
            $Name = $ThemeOrTarget
        }
    }

    if ($Name) {
        if ($Name -eq 'Custom' -and -not $NoSave) {
            Register-CustomThemeFromConfig -Force
        }
        $foundTheme = $null
        if ($script:Themes.TryGetValue($Name, [ref]$foundTheme)) {
            $script:ActiveThemeName = $Name
            $script:ActiveCache     = $foundTheme
            $global:DashboardConfig.State.CurrentTheme = $Name

            if (-not $NoSave -and $global:DashboardConfig.Config) {
                if (-not $global:DashboardConfig.Config.Options) { $global:DashboardConfig.Config.Options = @{} }
                $global:DashboardConfig.Config.Options.Theme = $Name
                if (Test-DashboardCommand 'WriteConfig') { WriteConfig }
            }
        }
    }

    $c = $script:ActiveCache
    if ($Target -and ($Target -is [System.Windows.Forms.Control])) {
        ApplyControlTheme $Target $c
    } else {
        $ui = $global:DashboardConfig.UI
        if ($ui) {
            foreach ($form in @($ui.MainForm, $ui.SettingsForm, $ui.ExtraForm)) {
                if ($form -and ($form -is [System.Windows.Forms.Control]) -and -not $form.IsDisposed) {
                    ApplyControlTheme $form $c
                }
            }
        }
        if ($global:DashboardConfig.Resources.OpenThemeableForms) {
            $openForms = @($global:DashboardConfig.Resources.OpenThemeableForms.ToArray())
            foreach ($of in $openForms) {
                if ($of -and -not $of.IsDisposed -and ($of -is [System.Windows.Forms.Control])) {
                    ApplyControlTheme $of $c
                }
            }
        }
        if (Test-DashboardCommand 'Update-NotificationTheme') {
            Update-NotificationTheme
        }
    }

    return $true
}

function Get-ThemeColors([string]$ThemeName) {
    if ([string]::IsNullOrWhiteSpace($ThemeName)) { return $script:ActiveCache }
    if ($ThemeName -eq 'Custom') {
        Register-CustomThemeFromConfig
    }
    $found = $null
    if ($script:Themes.TryGetValue($ThemeName, [ref]$found)) {
        return $found
    }
    return $script:ActiveCache
}

Export-ModuleMember -Function *