<# guide.psm1 #>

#region State and Configuration

$script:GuideState = @{
    CurrentStep = 0
    IsActive    = $false
    BubbleForm  = $null
    Language    = "EN"
}

$script:ButtonLabels = @{
    EN = @{ Prev = "Prev"; Next = "Next"; Close = "Close"; Finish = "Finish" }
    DE = @{ Prev = "Zur$([char]0x00FC)ck"; Next = "Weiter"; Close = "Schlie$([char]0x00DF)en"; Finish = "Fertig" }
    PT = @{ Prev = "Ant."; Next = "Prox."; Close = "Fechar"; Finish = "Fim" }
    PHL = @{ Prev = "Bumalik"; Next = "Sunod"; Close = "Isara"; Finish = "Tapos" }
}

#region Multi-Language Content
$script:GuideContent = @{
    EN = @(
        @{ Title = "Welcome"; Text = "Welcome to Entropia Dashboard! This interactive guide will walk you through the setup, features, and workflow.`n`nClick 'Next' to get started."; Target = "TitleLabel"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Settings"; Text = "First, we need to configure your game paths and preferences.`n`nI'll open the Settings menu for you now."; Target = "Settings"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Main Launcher"; Text = "Step 1: Select your game's 'Launcher.exe'.`n`nUse the 'Browse' button to find it. This path is used as the base for all profiles."; Target = "InputLauncher"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Process Name"; Text = "Usually 'neuz', but if you play on a server with a custom process name, enter it here so the dashboard can detect running clients."; Target = "InputProcess"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Max Clients"; Text = "This caps how many new clients a Launch starts for the selected profile.`n`nAlready-running clients do not count toward this limit."; Target = "InputMax"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Launcher Timeout"; Text = "How long (in seconds) the dashboard waits for the game launcher to open the client before retrying or giving up."; Target = "InputLauncherTimeout"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Profiles & Junctions"; Text = "To run multiple clients with different settings, you can create 'Profiles'.`nClick 'Create' to make a lightweight copy (Junction) of your client."; Target = "StartJunction"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Copy Neuz.exe"; Text = "Important! When the game has updates, use your Main Launcher to patch.`n`nThen, use this button to copy the complete 'bin32' and 'bin64' folders from your main folder to the profiles selected in the list (all profiles if none is selected). Files that cannot be overwritten are skipped and reported."; Target = "CopyNeuz"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Profile List"; Text = "Your created profiles appear here.`n`n- Double-click a path to open the folder in Explorer.`n- Select the profiles (Copy Neuz.exe copies bin32/bin64 to all selected profiles).`n- Selecting a row makes that profile the target of the Launch button."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Auto-Start Clients"; Text = "Every client that logs into a profile is remembered here.`n`n- Reconnect?: Reconnect this single client if it disconnects.`n- Hide?: Hide this single client window from Taskbar/Alt+Tab.`n- Hide Profiles (the switch below the list): checked, every client that has no row yet is hidden from Taskbar/Alt+Tab while minimized; unchecked, it stays visible. A client's own Hide? tick always wins.`n- Tick Ftool?/Macro?/QuickCmd? to start that tool for this single client once it has logged in (never while launching, logging in or logged out). Macro windows are opened but stay idle.`n- Click a client name to select it for its profile (the Launch button follows it).`n- Add / Rename / Remove maintain the list by hand."; Target = "ClientGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Saved Setups"; Text = "You can save specific groups of clients as a 'Setup'.`n`nExample: Save a setup with 1 Tank and 1 Damage Dealer to launch/login them both with one click later."; Target = "SetupGrid"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Login Settings"; Text = "Now let's configure Auto-Login.`n`nSwitching to the Login Settings tab..."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Select Profile"; Text = "Choose which profile you want to configure login coordinates for.`n`n'Default' applies to the main launcher path."; Target = "LoginProfileSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Screen Resolution"; Text = "Select the resolution your game client should use.`n`nThis updates the 'neuz.ini' file for that profile automatically and all automatic coordinates follow this change."; Target = "ResolutionSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Coordinate Pickers"; Text = "This is the most important part!`n`nBy default the coordinates are calculated from the profile's resolution, so nothing has to be picked. Tick 'Manual Coordinates' first, then click 'Set', click the corresponding button in your game window and wait 3 seconds to save the coordinate."; Target = "btnPickServer1"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Client Configuration"; Text = "Here you define which character logs into which client window.`n`nRow 1 = 1st Client in the list.`nServer/Channel/Char = 1 (First), 2 (Second), etc."; Target = "LoginConfigGrid"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Save Settings"; Text = "We are done with the basic setup. Let's save and go to the Dashboard."; Target = "Save"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Advanced Launch Menu"; Text = "Right-Click: Opens the Advanced Menu.`n`nHere you can launch specific Profiles, specific amounts (e.g., 5x Clients), or your saved Setups."; Target = "Launch"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Client List"; Text = "Running clients appear here.`n`nDouble-Click a row to bring that window to the front."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Context Menu"; Text = "Right-Click a row to see options:`n`n- Position: Move/Resize windows.`n- Set Hotkey: Assign a key to bring this specific window to front."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Hotkey Manager"; Text = "Speaking of Hotkeys... let's see where you manage them.`n`nOpening Settings > Hotkeys..."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Hotkey List"; Text = "All hotkeys you assigned via the Dashboard Context Menu appear here.`n`nDouble-click a row to edit the key combination."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Unregister"; Text = "To remove a hotkey, select the row and click here."; Target = "UnregisterHotkey"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Back to Dashboard"; Text = "Let's close settings and look at the Tools."; Target = "Cancel"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Ftool"; Text = "Select clients in the list, then click 'Ftool'.`n`nLeft-Click: Opens the tool.`nRight-Click: Saves the Ftool settings."; Target = "Ftool"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Macro Helper"; Text = "Select clients, then click 'Macro'.`n`nLeft-Click: Opens the tool.`nRight-Click: Saves the Macro settings."; Target = "Macro"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Auto Login"; Text = "Click 'Login' to automatically log in the selected clients using the settings we configured earlier."; Target = "LoginButton"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Terminate / Stop"; Text = "Controls selected clients:`n`nLeft-Click: Kills the process (Force Close).`nRight-Click: Disconnects the TCP connection (Instant Logout)."; Target = "Terminate"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Extras"; Text = "The 'Extra' menu contains the World Boss System, Notes, and Notification history and other tools.`n`nRight-Click: Opens ChatCommander (Quick chat macros)."; Target = "Extra"; Form = "MainForm"; Context = "Main" },
        @{ Title = "World Boss Setup"; Text = "Enter your Access Code (from Discord) and Username here.`n`nUse code 'TEST' to try it out locally without connecting to the server."; Target = "InputUser"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Listener & Images"; Text = "- Listener: Receive toast notifications when bosses spawn.`n- Show Images: Toggle graphical buttons."; Target = "WorldBossListener"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Boss List"; Text = "Click a boss to report it as spawned (Alerts everyone!).`n`nUse the toggle switch next to a button to mute notifications for that specific boss."; Target = "ButtonPanel"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Notes & Timers"; Text = "You can save text notes, or set Timers and Reminders here.`n`nUseful for tracking buffs or event times."; Target = "NoteGrid"; Form = "ExtraForm"; Context = "Extra_Notes" },
        @{ Title = "Notifications"; Text = "The history of all dashboard notifications (Boss spawns, errors, timers) is saved here for review.`nDouble-click to show the notification again."; Target = "NotificationGrid"; Form = "ExtraForm"; Context = "Extra_Notifications" },
        @{ Title = "Market Search"; Text = "Type an item name and the dashboard fuzzy-matches every item it knows - even half a name works.`n`nDown / Up picks a name, Enter accepts it, Esc closes the list."; Target = "MarketSearchInput"; Form = "ExtraForm"; Context = "Extra_Market" },
        @{ Title = "Maps & Wiki"; Text = "Left-click the 'Map' button for the Node Maps (every field, dungeon & event map).`n`nRight-click it for the built-in Wiki with guides, drop lists, and more."; Target = "Wiki"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Online / Offline"; Text = "Toggle between Online (Live Web) and Offline (Local Cache).`n`nOffline is faster but requires you to download the content first."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Edit Mode"; Text = "Turn on 'Edit Mode' to modify offline pages.`n`nYou can fix errors or add your own guides locally."; Target = "chkEdit"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Sync Features"; Text = "When viewing offline, use Sync buttons to update content.`n`n- Sync Node: Updates selected page.`n- Smart Sync: Adds missing pages without overwriting changes."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Finished"; Text = "That's it! You're ready to go.`n`nPress 'Finish' to close this guide."; Target = "InfoForm"; Form = "MainForm"; Context = "Main" }
    )

    DE = @(
        @{ Title = "Willkommen"; Text = "Willkommen beim Entropia Dashboard! Dieser interaktive Guide f$([char]0x00FC)hrt dich durch die Einrichtung und Funktionen.`n`nKlicke auf 'Weiter', um zu beginnen."; Target = "TitleLabel"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Einstellungen"; Text = "Zuerst konfigurieren wir die Spielpfade.`n`nIch $([char]0x00F6)ffne das Einstellungsmen$([char]0x00FC) f$([char]0x00FC)r dich."; Target = "Settings"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Main Launcher"; Text = "Schritt 1: W$([char]0x00E4)hle deine 'Launcher.exe'.`n`nBenutze 'Browse', um sie zu finden. Dieser Pfad dient als Basis f$([char]0x00FC)r alle Profile."; Target = "InputLauncher"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Prozessname"; Text = "Meistens 'neuz', aber bei P-Servern mit eigenem Namen hier eintragen, damit das Dashboard laufende Clients erkennt."; Target = "InputProcess"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Max Clients"; Text = "Begrenzt, wie viele neue Clients ein Launch f$([char]0x00FC)r das ausgew$([char]0x00E4)hlte Profil startet.`n`nBereits laufende Clients z$([char]0x00E4)hlen nicht mit."; Target = "InputMax"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Launcher Timeout"; Text = "Wie lange (in Sekunden) das Dashboard wartet, bis der Launcher den Client $([char]0x00F6)ffnet."; Target = "InputLauncherTimeout"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Profile & Junctions"; Text = "F$([char]0x00FC)r mehrere Clients mit verschiedenen Einstellungen kannst du 'Profile' erstellen.`nKlicke 'Create', um eine kleine Kopie (Junction) zu erstellen."; Target = "StartJunction"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Neuz.exe kopieren"; Text = "Wichtig! Patche das Spiel $([char]0x00FC)ber den Main Launcher.`n`nNutze dann diesen Button, um die kompletten Ordner 'bin32' und 'bin64' deines Hauptclients in die oben ausgew$([char]0x00E4)hlten Profile zu kopieren (alle Profile, falls keines ausgew$([char]0x00E4)hlt). Dateien, die nicht $([char]0x00FC)berschrieben werden k$([char]0x00F6)nnen, werden $([char]0x00FC)bersprungen und gemeldet."; Target = "CopyNeuz"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Profil-Liste"; Text = "Deine erstellten Profile erscheinen hier.`n`n- Doppelklick auf den Pfad, um den Ordner zu $([char]0x00F6)ffnen.`n- Mehrere Profile ausw$([char]0x00E4)hlen (Copy Neuz.exe kopiert bin32/bin64 in alle ausgew$([char]0x00E4)hlten Profile).`n- Die Auswahl bestimmt das Profil des Launch-Buttons."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Auto-Start Clients"; Text = "Jeder Client, der sich in einem Profil einloggt, wird hier gemerkt.`n`n- Reconnect?: Nur diesen Client bei Disconnect reconnecten.`n- Hide?: Nur dieses Client-Fenster vor Taskleiste/Alt+Tab verstecken.`n- Hide Profiles (der Schalter unterhalb der Liste): gesetzt wird jedes Client-Fenster, das noch keine eigene Zeile hat, beim Minimieren vor Taskleiste/Alt+Tab versteckt; ohne Haken bleibt es sichtbar. Der eigene Hide?-Haken eines Clients gewinnt immer.`n- Setze einen Haken bei Ftool?/Macro?/QuickCmd?, um das Tool nur f$([char]0x00FC)r diesen Client zu starten, sobald er eingeloggt ist (nie beim Starten, Einloggen oder Ausloggen). Makro-Fenster werden ge$([char]0x00F6)ffnet, bleiben aber inaktiv.`n- Klicke auf einen Client-Namen, um ihn f$([char]0x00FC)r sein Profil auszuw$([char]0x00E4)hlen (der Launch-Button folgt ihm).`n- Add / Rename / Remove pflegen die Liste von Hand."; Target = "ClientGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Gespeicherte Setups"; Text = "Speichere Gruppen von Clients als 'Setup'.`n`nBeispiel: 1 Tank + 1 Damage Dealer speichern, um sp$([char]0x00E4)ter beide mit einem Klick zu starten."; Target = "SetupGrid"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Login Einstellungen"; Text = "Jetzt konfigurieren wir den Auto-Login.`n`nWechsle zum Login-Tab..."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Profilwahl"; Text = "W$([char]0x00E4)hle das Profil, f$([char]0x00FC)r das du Koordinaten setzen willst.`n`n'Default' gilt f$([char]0x00FC)r den Hauptpfad."; Target = "LoginProfileSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Aufl$([char]0x00F6)sung"; Text = "W$([char]0x00E4)hle die Spielaufl$([char]0x00F6)sung.`n`nDies aktualisiert die 'neuz.ini' f$([char]0x00FC)r das Profil automatisch, alle automatischen Koordinaten folgen der Auswahl."; Target = "ResolutionSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Koordinaten-Picker"; Text = "Das Wichtigste!`n`nStandardm$([char]0x00E4)$([char]0x00DF)ig werden die Koordinaten aus der Aufl$([char]0x00F6)sung des Profils berechnet und muss nicht gesetzt werden. Setze zuerst den Haken bei 'Manual Coordinates', dann klicke 'Set', danach im Spiel auf den entsprechenden Button und warte 3 Sekunden zum Speichern."; Target = "btnPickServer1"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Client Konfiguration"; Text = "Hier legst du fest, welcher Charakter in welchem Fenster einloggt.`n`nZeile 1 = 1. Client.`nServer/Channel/Char = 1 (Erster), 2 (Zweiter), usw."; Target = "LoginConfigGrid"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Speichern"; Text = "Basiseinrichtung fertig. Speichern wir und gehen zum Dashboard."; Target = "Save"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Start-Men$([char]0x00FC)"; Text = "Rechtsklick: $([char]0x00D6)ffnet das erweiterte Men$([char]0x00FC).`n`nHier kannst du spezifische Profile, Mengen (z.B. 5x Clients) oder Setups starten."; Target = "Launch"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Client Liste"; Text = "Laufende Clients erscheinen hier.`n`nDoppelklick bringt das Fenster in den Vordergrund."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Kontextmen$([char]0x00FC)"; Text = "Rechtsklick f$([char]0x00FC)r Optionen:`n`n- Position: Fenster verschieben.`n- Set Hotkey: Taste zuweisen, um dieses Fenster in den Vordergrund zu holen."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Hotkey Manager"; Text = "Wo wir gerade von Hotkeys sprechen...`n`n$([char]0x00D6)ffne Einstellungen > Hotkeys..."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Hotkey Liste"; Text = "Alle zugewiesenen Hotkeys erscheinen hier.`n`nDoppelklick zum Bearbeiten."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Entfernen"; Text = "W$([char]0x00E4)hle eine Zeile und klicke hier, um den Hotkey zu l$([char]0x00F6)schen."; Target = "UnregisterHotkey"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Zur$([char]0x00FC)ck"; Text = "Schlie$([char]0x00DF)en wir die Einstellungen und schauen uns die Tools an."; Target = "Cancel"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Ftool"; Text = "Clients ausw$([char]0x00E4)hlen, dann 'Ftool' klicken.`n`nLinksklick: $([char]0x00D6)ffnet das Tool.`nRechtsklick: Speichert die Ftool-Einstellungen."; Target = "Ftool"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Makro Helper"; Text = "Clients ausw$([char]0x00E4)hlen, dann 'Macro'.`n`nLinksklick: $([char]0x00D6)ffnet das Tool.`nRechtsklick: Speichert die  Makro-Einstellungen."; Target = "Macro"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Auto Login"; Text = "Klicke 'Login', um die ausgew$([char]0x00E4)hlten Clients automatisch einzuloggen."; Target = "LoginButton"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Beenden / Stop"; Text = "Steuert ausgew$([char]0x00E4)hlte Clients:`n`nLinksklick: Prozess beenden (Force Close).`nRechtsklick: TCP trennen (Instant Logout)."; Target = "Terminate"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Extras"; Text = "Das 'Extra' Men$([char]0x00FC) enth$([char]0x00E4)lt World Boss, Notizen und Historie und andere Tools.`n`nRechtsklick: $([char]0x00D6)ffnet ChatCommander (Chat-Makros)."; Target = "Extra"; Form = "MainForm"; Context = "Main" },
        @{ Title = "World Boss Setup"; Text = "Gib hier deinen Zugangscode (Discord) und Namen ein.`n`nNutze 'TEST', um es lokal zu probieren."; Target = "InputUser"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Listener & Bilder"; Text = "- Listener: Toast-Benachrichtigung bei Boss-Spawn.`n- Show Images: Grafische Buttons anzeigen."; Target = "WorldBossListener"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Boss Liste"; Text = "Klicke einen Boss, um ihn als gespawnt zu melden.`n`nNutze den Schalter daneben, um Benachrichtigungen f$([char]0x00FC)r diesen Boss stummzuschalten."; Target = "ButtonPanel"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Notizen & Timer"; Text = "Speichere Textnotizen oder setze Timer/Erinnerungen hier."; Target = "NoteGrid"; Form = "ExtraForm"; Context = "Extra_Notes" },
        @{ Title = "Benachrichtigungen"; Text = "Die Historie aller Dashboard-Meldungen wird hier gespeichert.`nDoppelklick zeigt die Meldung erneut an."; Target = "NotificationGrid"; Form = "ExtraForm"; Context = "Extra_Notifications" },
        @{ Title = "Markt Suche"; Text = "Gib einen Gegenstandsnamen ein und das Dashboard gleicht alles ab, was es kennt - auch ein halber Name reicht.`n`nUnten / Oben $([char]0x00E4)hlt einen Namen, Enter $([char]0x00FC)bernimmt ihn, Esc schlie$([char]0x00DF)t die Liste."; Target = "MarketSearchInput"; Form = "ExtraForm"; Context = "Extra_Market" },
        @{ Title = "Maps & Wiki"; Text = "Linksklick auf 'Map' zeigt die Node Maps (alle Felder, Dungeons & Event-Maps).`n`nRechtsklick zeigt das integrierte Wiki mit Guides, Droplisten und mehr."; Target = "Wiki"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Online / Offline"; Text = "Wechsle zwischen Online (Live Web) und Offline (Lokaler Cache).`n`nOffline ist schneller, erfordert aber vorherigen Download."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Bearbeitungsmodus"; Text = "Aktiviere 'Edit Mode', um Offline-Seiten zu $([char]0x00E4)ndern.`n`nDu kannst Fehler beheben oder eigene Guides schreiben."; Target = "chkEdit"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Sync Funktionen"; Text = "Offline-Modus Sync-Optionen:`n`n- Sync Node: Aktualisiert die Seite.`n- Smart Sync: F$([char]0x00FC)gt fehlende Seiten hinzu, ohne $([char]0x00C4)nderungen zu $([char]0x00FC)berschreiben."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Fertig"; Text = "Das war's! Du bist bereit.`n`nDr$([char]0x00FC)cke 'Fertig', um den Guide zu schlie$([char]0x00DF)en."; Target = "InfoForm"; Form = "MainForm"; Context = "Main" }
    )

    PT = @(
        @{ Title = "Bem-vindo"; Text = "Bem-vindo ao Entropia Dashboard! Este guia interativo mostrar$([char]0x00E1) a configura$([char]0x00E7)$([char]0x00E3)o e recursos.`n`nClique em 'Prox.' para come$([char]0x00E7)ar."; Target = "TitleLabel"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Configura$([char]0x00E7)$([char]0x00F5)es"; Text = "Primeiro, precisamos configurar os caminhos do jogo.`n`nVou abrir o menu de Configura$([char]0x00E7)$([char]0x00F5)es para voc$([char]0x00EA)."; Target = "Settings"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Main Launcher"; Text = "Passo 1: Selecione o 'Launcher.exe' do jogo.`n`nUse 'Browse' para encontrar. Este caminho $([char]0x00E9) a base para todos os perfis."; Target = "InputLauncher"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Process Name"; Text = "Geralmente 'neuz'. Se jogar em servidor com nome customizado, digite aqui para detectar os clientes."; Target = "InputProcess"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Max Clients"; Text = "Limita quantos clientes novos um Launch abre para o perfil selecionado.`n`nClientes j$([char]0x00E1) em execu$([char]0x00E7)$([char]0x00E3)o n$([char]0x00E3)o contam."; Target = "InputMax"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Launcher Timeout"; Text = "Tempo (em segundos) que o dashboard espera o launcher abrir o jogo antes de tentar novamente."; Target = "InputLauncherTimeout"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Perfis e Jun$([char]0x00E7)$([char]0x00F5)es"; Text = "Para rodar v$([char]0x00E1)rios clientes com configura$([char]0x00E7)$([char]0x00F5)es diferentes, crie 'Profiles'.`nClique em 'Create' para fazer uma c$([char]0x00F3)pia leve (Junction)."; Target = "StartJunction"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Copy Neuz.exe"; Text = "Importante! Quando houver update, use o Launcher Principal para atualizar.`n`nDepois, use este bot$([char]0x00E3)o para copiar as pastas completas 'bin32' e 'bin64' do seu cliente principal para os perfis selecionados acima (todos os perfis, se nenhum estiver selecionado). Arquivos que n$([char]0x00E3)o puderem ser sobrescritos ser$([char]0x00E3)o ignorados e reportados."; Target = "CopyNeuz"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Lista de Perfis"; Text = "Seus perfis aparecem aqui.`n`n- D$([char]0x00FA)plo clique em um caminho para abrir a pasta no Explorer.`n- Selecione os perfis (o Copy Neuz.exe copia bin32/bin64 para todos os perfis selecionados).`n- Selecionar uma linha define o perfil do bot$([char]0x00E3)o Launch."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Auto-Start Clients"; Text = "Todos os clientes que entram em um perfil s$([char]0x00E3)o lembrados aqui.`n`n- Reconnect?: Reconecta apenas este cliente se ele cair.`n- Hide?: Esconde apenas esta janela de cliente da Barra de Tarefas/Alt+Tab.`n- Hide Profiles (a op$([char]0x00E7)$([char]0x00E3)o abaixo da lista): marcada, todo cliente que ainda n$([char]0x00E3)o tem linha pr$([char]0x00F3)pria $([char]0x00E9) ocultado da barra de tarefas/Alt+Tab ao minimizar; desmarcada, ele continua vis$([char]0x00ED)vel. A marca Hide? do cliente sempre prevalece.`n- Marque Ftool?/Macro?/QuickCmd? para abrir essa ferramenta apenas para este cliente assim que ele fizer login (nunca durante o lan$([char]0x00E7)amento, login ou logout). As janelas de Macro s$([char]0x00E3)o abertas, mas ficam paradas.`n- Ao clicar no nome de um cliente, ele vira o cliente selecionado do perfil (o Launch segue ele).`n- Add / Rename / Remove mant$([char]0x00E9)m a lista manualmente."; Target = "ClientGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Saved Setups"; Text = "Salve grupos de clientes como um 'Setup'.`n`nExemplo: Salve 1 Tank e 1 DPS para logar ambos com um clique depois."; Target = "SetupGrid"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Config de Login"; Text = "Agora vamos configurar o Auto-Login.`n`nTrocando para a aba Login Settings..."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Selecionar Perfil"; Text = "Escolha qual perfil configurar.`n`n'Default' aplica-se ao caminho principal."; Target = "LoginProfileSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Resolu$([char]0x00E7)$([char]0x00E3)o"; Text = "Selecione a resolu$([char]0x00E7)$([char]0x00E3)o do jogo.`n`nIsso atualiza o 'neuz.ini' do perfil automaticamente e todas as coordenadas autom$([char]0x00E1)ticas acompanham a mudan$([char]0x00E7)a."; Target = "ResolutionSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Captura de Coordenadas"; Text = "A parte mais importante!`n`nPor padr$([char]0x00E3)o as coordenadas s$([char]0x00E3)o calculadas pela resolu$([char]0x00E7)$([char]0x00E3)o do perfil e n$([char]0x00E3)o precisam ser definidas. Marque primeiro 'Manual Coordinates', depois clique em 'Set', clique no bot$([char]0x00E3)o correspondente na janela do jogo e espere 3 segundos."; Target = "btnPickServer1"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Config do Cliente"; Text = "Defina qual personagem loga em qual janela.`n`nLinha 1 = 1$([char]0x00BA) Cliente.`nServer/Channel/Char = 1 (Primeiro), 2 (Segundo), etc." ; Target = "LoginConfigGrid"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Salvar"; Text = "Configura$([char]0x00E7)$([char]0x00E3)o b$([char]0x00E1)sica pronta. Vamos salvar e ir ao Dashboard."; Target = "Save"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Menu de Launch"; Text = "Clique Direito: Abre Menu Avan$([char]0x00E7)ado.`n`nAqui voc$([char]0x00EA) pode abrir perfis espec$([char]0x00ED)ficos, quantidades especificas (ex: 5x) ou seus Setups salvos."; Target = "Launch"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Lista de Cliente"; Text = "Clients rodando aparecem aqui.`n`nClique Duplo traz a janela para frente."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Menu de Contexto"; Text = "Clique Direito em um cliente para op$([char]0x00E7)$([char]0x00F5)es:`n`n- Position: Mover/redimensionar janela.`n- Set Hotkey: Atribuir atalho para trazer esta janela para frente."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Gerenciar Atalhos"; Text = "Falando em atalhos... veja onde gerenci$([char]0x00E1)-las.`n`nAbrindo Settings > Hotkeys..."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Lista de Atalhos"; Text = "Todos os atalhos atribu$([char]0x00ED)dos aparecem aqui.`n`nClique duplo para editar."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Remover atalhos"; Text = "Para remover atalhos, selecione a linha e clique aqui."; Target = "UnregisterHotkey"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "De volta ao Dashboard"; Text = "Vamos fechar as configura$([char]0x00E7)$([char]0x00F5)es e ver as Ferramentas."; Target = "Cancel"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Ftool"; Text = "Selecione clientes na lista, clique em 'Ftool'.`n`nClique esquerdo: Abre a ferramenta.`nClique direito: Save todas as configura$([char]0x00E7)$([char]0x00F5)es do Ftool."; Target = "Ftool"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Macro Helper"; Text = "Selecione clientes, clique em 'Macro'.`n`nClique esquerdo: Abre a ferramenta.`nClique direito: Save todas as configura$([char]0x00E7)$([char]0x00F5)es do Macro."; Target = "Macro"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Auto Login"; Text = "Clique 'Login' para logar automaticamente os clientes selecionados utilizando as configura$([char]0x00E7)$([char]0x00F5)es que fizemos mais cedo."; Target = "LoginButton"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Terminar / Parar"; Text = "Controla clientes selecionados:`n`nClique esquerdo: Fecha processo (Force Close).`nClique direito: Desconecta TCP (Instant Logout)."; Target = "Terminate"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Extras"; Text = "O menu 'Extra' cont$([char]0x00E9)m World Boss, Notas e Hist$([char]0x00F3)rico e outras extens$([char]0x00F5)es.`n`nClique direito: Abre ChatCommander (Macros de chat)."; Target = "Extra"; Form = "MainForm"; Context = "Main" },
        @{ Title = "World Boss Setup"; Text = "Insira seu C$([char]0x00F3)digo de Acesso (Discord) e Nome.`n`nUse 'TEST' para testar localmente."; Target = "InputUser"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Listener e Imagens"; Text = "- Listener: Notifica$([char]0x00E7)$([char]0x00E3)o quando boss nascer.`n- Show Images: Bot$([char]0x00F5)es gr$([char]0x00E1)ficos."; Target = "WorldBossListener"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Lista de Boss"; Text = "Clique num boss para reportar nascimento.`n`nUse a chave ao lado para silenciar notifica$([char]0x00E7)$([char]0x00F5)es daquele boss."; Target = "ButtonPanel"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Notas e Timers"; Text = "Salve notas de texto ou crie Timers/Lembretes aqui."; Target = "NoteGrid"; Form = "ExtraForm"; Context = "Extra_Notes" },
        @{ Title = "Notifica$([char]0x00E7)$([char]0x00F5)es"; Text = "O hist$([char]0x00F3)rico de notifica$([char]0x00E7)$([char]0x00F5)es fica salvo aqui.`nClique duplo para ver a mensagem novamente."; Target = "NotificationGrid"; Form = "ExtraForm"; Context = "Extra_Notifications" },
        @{ Title = "Pesquisa do Mercado"; Text = "Digite o nome de um item e o dashboard compara com tudo o que conhece - at$([char]0x00E9) meio nome serve.`n`nBaixo / Cima escolhe um nome, Enter confirma, Esc fecha a lista."; Target = "MarketSearchInput"; Form = "ExtraForm"; Context = "Extra_Market" },
        @{ Title = "Mapas e Wiki"; Text = "Clique com o bot$([char]0x00E3)o esquerdo em 'Map' para os Mapas de Nodes (todos os campos, dungeons e mapas de eventos).`n`nClique com o bot$([char]0x00E3)o direito para a Wiki integrada com guias, listas de drops e mais."; Target = "Wiki"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Online / Offline"; Text = "Alterne entre Online (Web) e Offline (Cache Local).`n`nOffline $([char]0x00E9) mais r$([char]0x00E1)pido, mas requer download do conte$([char]0x00FA)do."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Modo de Edi$([char]0x00E7)$([char]0x00E3)o"; Text = "Ative o 'Edit Mode' para modificar p$([char]0x00E1)ginas offline.`n`nVoc$([char]0x00EA) pode corrigir erros ou adicionar seus pr$([char]0x00F3)prios guias."; Target = "chkEdit"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Sincroniza$([char]0x00E7)$([char]0x00E3)o"; Text = "No modo offline, use bot$([char]0x00F5)es de Sync:`n`n- Sync Node: Atualiza p$([char]0x00E1)gina selecionada.`n- Smart Sync: Adiciona p$([char]0x00E1)ginas faltantes sem sobrescrever."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Fim"; Text = "$([char]0x00C9) isso! Voc$([char]0x00EA) est$([char]0x00E1) pronto.`n`nPressione 'Fim' para fechar o guia."; Target = "InfoForm"; Form = "MainForm"; Context = "Main" }
    )

    PHL = @(
        @{ Title = "Welcome"; Text = "Welcome sa Entropia Dashboard! Gagabayan ka ng interactive na gabay na ito sa pag-setup, features, at workflow.`n`nClick 'Next' para makapagumpisa."; Target = "TitleLabel"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Settings"; Text = "Una, kailangan naming i-configure ang iyong mga path ng laro at mga kagustuhan.`n`nI'll buksan ang menu ng Mga Setting para sa iyo ngayon."; Target = "Settings"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Main Launcher"; Text = "Step 1: Piliin ang 'Launcher.exe ng iyong laro'.`n`nUse ang 'Browse' na button upang mahanap ito. Ang path na ito ay ginagamit bilang batayan para sa lahat ng mga profile."; Target = "InputLauncher"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Process Name"; Text = "Karaniwang 'neuz', ngunit kung naglalaro ka sa isang server na may custom na pangalan ng proseso, ilagay ito dito para matukoy ng dashboard ang mga tumatakbong kliyente."; Target = "InputProcess"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Max Clients"; Text = "Nililimitahan nito kung ilang bagong client ang ilulunsad ng Launch para sa napiling profile. Hindi binibilang ang mga client na tumatakbo na."; Target = "InputMax"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Launcher Timeout"; Text = "Gaano katagal (in seconds) naghihintay ang dashboard para buksan ng launcher ng laro ang client bago muling subukan o magsara."; Target = "InputLauncherTimeout"; Form = "SettingsForm"; Context = "Settings_Launch" },
        @{ Title = "Profiles & Junctions"; Text = "Upang magpatakbo ng maraming clients na may iba't ibang setting, maaari kang lumikha ng 'Mga Profile'.` I-click ang 'Create' upang makagawa ng magaan na kopya (Junction) ng iyong client."; Target = "StartJunction"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Copy Neuz.exe"; Text = "Mahalaga! Kapag may mga update ang laro, gamitin ang iyong Main Launcher para mag-patch. Pagkatapos, gamitin ang button na ito para kopyahin ang buong 'bin32' at 'bin64' folder ng iyong main client sa mga profile na naka-select sa itaas (lahat ng profile kung walang naka-select). Ang mga file na hindi ma-overwrite ay ini-skip at ini-report."; Target = "CopyNeuz"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Profile List"; Text = "Lalabas dito ang iyong mga ginawang profile. Doble-klik sa path para buksan ang folder sa Explorer. Piliin ang mga profile (kinokopya ng Copy Neuz.exe ang bin32/bin64 sa lahat ng naka-select na profile). Ang pag-click sa row ang magiging profile ng Launch button."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Auto-Start Clients"; Text = "Bawat client na naglo-login sa isang profile ay naaalala dito. Reconnect?: Ire-reconnect lang ang client na ito kapag na-disconnect. Hide?: Itatago lang ang window ng client na ito sa Taskbar/Alt+Tab. 'Hide Profiles': kapag naka-check ang switch sa ilalim ng listahan, itatago sa Taskbar/Alt+Tab habang naka-minimize ang bawat client na wala pang sariling row; kapag hindi naka-check, mananatili itong nakikita. Palaging panalo ang sariling Hide? check ng client. I-check ang Ftool?/Macro?/QuickCmd? para buksan ang tool ng partikular na client na ito kapag nakalogin na (hindi habang nagla-launch, naglo-login, o naka-logout). Binubuksan ang mga Macro window ngunit nananatiling idle. Ang pag-click sa pangalan ng client ang magiging selected client ng profile (susundin ito ng Launch button). Ang Add / Rename / Remove ang bahala sa pagpapanatili ng listahan."; Target = "ClientGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Saved Setups"; Text = "Maaari mong i-save ang mga partikular na grupo ng mga client bilang isang 'Setup'. Halimbawa: Mag-save ng setup na may 1 Tank at 1 Damage Dealer para ilunsad/i-login silang pareho sa isang click mamaya."; Target = "SetupGrid"; Form = "SettingsForm"; Context = "Settings_Launch" },
		@{ Title = "Login Settings"; Text = "Ngayon ay i-configure natin ang Auto-Login. Paglipat sa tab na Mga Setting ng Pag-login..."; Target = "ProfileGrid"; Form = "SettingsForm"; Context = "Settings_General" },
        @{ Title = "Select Profile"; Text = "Piliin kung aling profile ang gusto mong i-configure ang mga coordinate sa pag-log in. 'Default' ang pangunahing path ng launcher."; Target = "LoginProfileSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
		@{ Title = "Screen Resolution"; Text = "Piliin ang resolution na dapat gamitin ng iyong game client. Automatically ina-update nito ang 'neuz.ini' file para sa profile na iyon at sinusundan ng lahat ng automatic na coordinate ang pagbabago."; Target = "ResolutionSelector"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Coordinate Pickers"; Text = "Ito ang pinakamahalagang bahagi! Ayon sa default ay kinakalkula ang mga coordinate mula sa resolution ng profile kaya hindi na kailangang mag-set. I-check muna ang 'Manual Coordinates', pagkatapos i-click ang 'Set', i-click ang kaukulang button sa window ng iyong laro at maghintay ng 3 seconds upang i-save ang coordinate."; Target = "btnPickServer1"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Client Configuration"; Text = "Dito mo tutukuyin kung aling character ang magla-log in kung aling window ng client. Row 1 = 1st Client sa listahan. Server/Channel/Char = 1 (Una), 2 (Pangalawa), etc."; Target = "LoginConfigGrid"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Save Settings"; Text = "Tapos na ang basic setup. Mag-save tayo at pumunta sa Dashboard."; Target = "Save"; Form = "SettingsForm"; Context = "Settings_Login" },
        @{ Title = "Advanced Launch Menu"; Text = "I-right-click: Binubuksan ang Advanced na Menu. Dito maaari kang maglunsad ng mga partikular na Profile, mga partikular na amounts (hal., 5x na Clients), o ang iyong mga naka-save na Setup."; Target = "Launch"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Client List"; Text = "Lalabas dito ang mga Running clients. Double-I-click ang isang row para dalhin ang window na iyon sa harap."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Context Menu"; Text = "Mag-right-click sa isang row para makita ang mga option: Position: Ilipat/Baguhin ang laki ng mga window. Mag-Set ng Hotkey: Magtalaga ng key upang dalhin ang partikular na window na ito sa harap."; Target = "DataGridMain"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Hotkey Manager"; Text = "Speaking of Hotkeys... tingnan natin kung saan mo maimamanage ang mga ito. Open Setting > Hotkey..."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Hotkey List"; Text = "Lalabas dito ang lahat ng hotkey na itinalaga mo sa pamamagitan ng Menu ng Dashboard. I-double-click ang isang row para i-edit ang combination ng key."; Target = "HotkeysGrid"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Unregister"; Text = "Upang alisin ang isang hotkey, piliin ang row at mag-click dito."; Target = "UnregisterHotkey"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Back to Dashboard"; Text = "Isara natin ang mga setting at tingnan ang Tools."; Target = "Cancel"; Form = "SettingsForm"; Context = "Settings_Hotkeys" },
        @{ Title = "Ftool"; Text = "Pumili ng mga Client sa listahan, pagkatapos ay i-click ang 'Ftool'.`n`nLeft-Click: Bubuksan ang tool.`nRight-Click: Save ang lahat ng setting ng Ftool."; Target = "Ftool"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Macro Helper"; Text = "Pumili ng mga kliyente, pagkatapos ay i-click ang 'Macro'.`n`nLeft-Click: Bubuksan ang tool.`nRight-Click: Save ang lahat ng setting ng Macro."; Target = "Macro"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Auto Login"; Text = "I-click ang 'Login' para automatically mag-log in sa mga napiling client gamit ang mga setting na na-configure kanina."; Target = "LoginButton"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Terminate / Stop"; Text = "Kinokontrol ang mga napiling kliyente: Left-Click: Pinapatay ang proseso (Force Close). Right-Click: Disconnect ang TCP connection (Instant Logout)."; Target = "Terminate"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Extras"; Text = "Ang 'Extra' menu ay naglalaman ng World Boss, Notes, at History, extra tools.`n`nRight-Click: Bubuksan ang ChatCommander (Chat Macros)."; Target = "Extra"; Form = "MainForm"; Context = "Main" },
        @{ Title = "World Boss Setup"; Text = "Ilagay ang iyong Access Code (mula sa Discord) at Username dito. Gamitin ang code na 'TEST' upang subukan ito nang lokal nang hindi kumokonekta sa server."; Target = "InputUser"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Listener & Images"; Text = "- Listener: Makatanggap ng mga notification kapag nag-spawn ang mga boss. Ipakita ang Mga Larawan: I-toggle ang mga graphical na button."; Target = "WorldBossListener"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Boss List"; Text = "I-click ang isang boss para ma-report ito bilang spawned (Alerts everyone!). Gamitin ang toggle switch sa tabi ng isang button para i-mute ang mga notification para sa partikular na boss na iyon."; Target = "ButtonPanel"; Form = "ExtraForm"; Context = "Extra_WorldBoss" },
        @{ Title = "Notes & Timers"; Text = "Maaari kang mag-save ng mga text, notes, o mag-set ng mga Timer at Paalala dito. Kapaki-pakinabang para sa pagsubaybay sa mga buff o oras ng kaganapan."; Target = "NoteGrid"; Form = "ExtraForm"; Context = "Extra_Notes" },
        @{ Title = "Notifications"; Text = "Ang History ng lahat ng notification sa dashboard (Boss spawns, errors, timers) ay naka-save dito para sa pagsusuri. Double-click upang ipakita muli ang notification."; Target = "NotificationGrid"; Form = "ExtraForm"; Context = "Extra_Notifications" },
        @{ Title = "Paghahanap sa Market"; Text = "Mag-type ng item name at itutugma ng dashboard ang lahat ng kilala nito - kahit kalahating name ayos na.`n`nDown / Up ang pumipili ng name, Enter ang tumatanggap, Esc ang nagsasara ng listahan."; Target = "MarketSearchInput"; Form = "ExtraForm"; Context = "Extra_Market" },
        @{ Title = "Maps & Wiki"; Text = "Left-click ang 'Map' para sa Node Maps (lahat ng field, dungeon at event map).`n`nRight-click para sa built-in Wiki na may guides, drop lists, at iba pa."; Target = "Wiki"; Form = "MainForm"; Context = "Main" },
        @{ Title = "Online / Offline"; Text = "Mag-switch sa pagitan ng Online (Live Web) at Offline (Local Cache).`n`nMas mabilis ang Offline pero kailangan mong i-download muna ang content."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Edit Mode"; Text = "I-on ang 'Edit Mode' para baguhin ang offline pages.`n`nPwede kang mag-fix ng errors o magdagdag ng sariling guides."; Target = "chkEdit"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Sync Features"; Text = "Kapag offline, gamitin ang Sync buttons:`n`n- Sync Node: I-update ang napiling page.`n- Smart Sync: Magdagdag ng missing pages nang hindi binubura ang changes."; Target = "chkViewMode"; Form = "WikiForm"; Context = "Wiki" },
        @{ Title = "Finished"; Text = "Yun lang! Lahat ay nakasetup na. Pindutin ang 'Tapos' upang isara ang guide na ito."; Target = "InfoForm"; Form = "MainForm"; Context = "Main" }
    )
}
#endregion

#endregion

#region Core Functions

function Show-Guide {
	
    if ($script:GuideState.IsActive) {
        if ($script:GuideState.BubbleForm) {
            $script:GuideState.BubbleForm.BringToFront()
            return
        }
    }

    $script:GuideState.CurrentStep = 0
    $script:GuideState.IsActive = $true

    Show-GuideBubble
}

function Close-Guide {
    if ($script:GuideState.BubbleForm) {
        $script:GuideState.BubbleForm.Close()
        $script:GuideState.BubbleForm.Dispose()
        $script:GuideState.BubbleForm = $null
	}
    if ($script:GuideState.BubbleForm -and -not $script:GuideState.BubbleForm.IsDisposed) {
        try { $script:GuideState.BubbleForm.Close() } catch {}
        try { $script:GuideState.BubbleForm.Dispose() } catch {}
    }
    $script:GuideState.BubbleForm = $null
    $script:GuideState.IsActive = $false
    Write-Verbose "Guide Closed"
}

function Show-GuideBubble {
    $stepIndex = $script:GuideState.CurrentStep
    $lang = $script:GuideState.Language
    $steps = $script:GuideContent[$lang]

    if ($stepIndex -lt 0 -or $stepIndex -ge $steps.Count) {
        Close-Guide
        return
    }

    $step = $steps[$stepIndex]
    
    Switch-GuideContext -Context $step.Context

    $targetControl = Resolve-TargetControl -FormName $step.Form -ControlName $step.Target
    
    if (-not $targetControl -or $targetControl.IsDisposed) {
        Write-Warning "Guide: Could not find target control '$($step.Target)' on form '$($step.Form)'. Skipping to next step."
        $script:GuideState.CurrentStep++
        Show-GuideBubble
        return
    }

    if (-not $script:GuideState.BubbleForm -or $script:GuideState.BubbleForm.IsDisposed) {
        $script:GuideState.BubbleForm = CreateBubbleForm
        Register-OpenForm $script:GuideState.BubbleForm
    }
    
    $form = $script:GuideState.BubbleForm
    Update-BubbleContent -Form $form -Step $step -StepIndex $stepIndex -TotalSteps $steps.Count
    
    PositionBubbleForm -Form $form -Target $targetControl

    if (-not $form.Visible) {
        $form.Show()
    }
    $form.BringToFront()
}

function Select-SettingsTab {
    param([string]$Name)

    $UI = $global:DashboardConfig.UI
    if (-not $UI -or -not $UI.SettingsTabs) { return }

    foreach ($page in $UI.SettingsTabs.TabPages) {
        if ($page.Text -eq $Name) {
            $UI.SettingsTabs.SelectedTab = $page
            return
        }
    }
}

function Select-ExtraTab {
    param([string]$Name, [string]$SubName)

    $UI = $global:DashboardConfig.UI
    if (-not $UI -or -not $UI.ExtraForm) { return }

    if (-not $UI.ExtraForm.Visible) { ShowExtraForm }
    $UI.ExtraForm.Activate()

    $tc = $UI.ExtraForm.Controls | Where-Object { $_.GetType().Name -match 'TabControl' } | Select-Object -First 1
    if (-not $tc) { return }

    $page = $null
    foreach ($candidate in $tc.TabPages) {
        if ($candidate.Text -eq $Name) { $page = $candidate; break }
    }
    if (-not $page) { return }
    $tc.SelectedTab = $page

    if (-not $SubName) { return }

    $inner = $page.Controls | Where-Object { $_.GetType().Name -match 'TabControl' } | Select-Object -First 1
    if (-not $inner) { return }

    foreach ($subPage in $inner.TabPages) {
        if ($subPage.Text -eq $SubName) {
            $inner.SelectedTab = $subPage
            return
        }
    }
}

function Switch-GuideContext {
    param([string]$Context)

    $UI = $global:DashboardConfig.UI

    switch ($Context) {
        "Main" {
            if ($UI.SettingsForm.Visible) { HideSettingsForm }
            if ($UI.ExtraForm.Visible) { HideExtraForm }
            $wForm = if ($UI.WikiForm) { $UI.WikiForm } else { $global:DashboardConfig.Resources.WikiForm }
            if ($wForm -and -not $wForm.IsDisposed -and $wForm.Visible) { $wForm.Close() }
            $UI.MainForm.Activate()
        }
        "Settings_General" {
            if (-not $UI.SettingsForm.Visible) { ShowSettingsForm }
            Select-SettingsTab -Name 'General'
            $UI.SettingsForm.Activate()
        }
        "Settings_Launch" {
            if (-not $UI.SettingsForm.Visible) { ShowSettingsForm }
            Select-SettingsTab -Name 'Launch Settings'
            $UI.SettingsForm.Activate()
        }
        "Settings_Login" {
            if (-not $UI.SettingsForm.Visible) { ShowSettingsForm }
            Select-SettingsTab -Name 'Login Settings'
            $UI.SettingsForm.Activate()
        }
        "Settings_Hotkeys" {
            if (-not $UI.SettingsForm.Visible) { ShowSettingsForm }
            Select-SettingsTab -Name 'Hotkeys'
            $UI.SettingsForm.Activate()
        }
        "Extra" {
            if (-not $UI.ExtraForm.Visible) { ShowExtraForm }
            $UI.ExtraForm.Activate()
        }
        "Extra_WorldBoss" {
            Select-ExtraTab -Name 'Tools' -SubName 'World Boss'
        }
        "Extra_Notes" {
            Select-ExtraTab -Name 'Notes'
        }
        "Extra_Notifications" {
            Select-ExtraTab -Name 'Notifications'
        }
        "Extra_Market" {
            Select-ExtraTab -Name 'Market'
        }
        "Wiki" {
            $wForm = if ($UI.WikiForm) { $UI.WikiForm } else { $global:DashboardConfig.Resources.WikiForm }
            if (-not $wForm -or -not $wForm.Visible) { Show-Wiki }
            $wForm = if ($UI.WikiForm) { $UI.WikiForm } else { $global:DashboardConfig.Resources.WikiForm }
            if ($wForm) { $wForm.Activate() }
        }
    }
}

function Resolve-TargetControl {
    param($FormName, $ControlName)
    
    $UI = $global:DashboardConfig.UI
    if (-not $UI) { return $null }

    if ($ControlName -match '^btnPick(.+)$') {
        $key = $Matches[1]
        if ($UI.LoginPickers -and $UI.LoginPickers.ContainsKey($key)) {
            return $UI.LoginPickers[$key].Button
        }
    }

    if ($UI | Get-Member -Name $ControlName) {
        return $UI.$ControlName
    }

    $form = $null
    if ($FormName -eq 'MainForm') { $form = $UI.MainForm }
    elseif ($FormName -eq 'SettingsForm') { $form = $UI.SettingsForm }
    elseif ($FormName -eq 'ExtraForm') { $form = $UI.ExtraForm }
    elseif ($FormName -eq 'WikiForm') { 
        $form = if ($UI.WikiForm) { $UI.WikiForm } else { $global:DashboardConfig.Resources.WikiForm }
    }
    
    if ($form) {
        $foundControls = $form.Controls.Find($ControlName, $true)
        if ($foundControls.Count -gt 0) {
            return $foundControls[0]
        }
    }

    return $null
}

function CreateBubbleForm {
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.StartPosition = 'Manual'
    $form.Width = 300
    $form.Height = 160
    $form.TopMost = $true
    $form.BackColor = Get-Theme 'TertiaryBg'
    $form.ShowInTaskbar = $false
    
    $form.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen((Get-Theme 'TopBarAccent'), 2)
        try {
            $rect = $s.ClientRectangle
            $rect.Width -= 1
            $rect.Height -= 1
            $e.Graphics.DrawRectangle($pen, $rect)
        } finally {
            $pen.Dispose()
        }
    })

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblTitle.Size = New-Object System.Drawing.Size(220, 25)
    $lblTitle.Font = Get-UIFont 'Segoe UI' 11 'Bold'
    $lblTitle.ForeColor = Get-Theme 'WhiteText'
    $lblTitle.Name = "lblTitle"
    $form.Controls.Add($lblTitle)

    $cbLang = New-Object Custom.DarkComboBox
    $cbLang.Location = New-Object System.Drawing.Point(235, 8)
    $cbLang.Size = New-Object System.Drawing.Size(55, 23)
    $cbLang.DropDownStyle = 'DropDownList'
    $cbLang.FlatStyle = 'Flat'
    $cbLang.BackColor = Get-Theme 'InputBgAlt'
    $cbLang.ForeColor = Get-Theme 'PrimaryText'
    $cbLang.Font = Get-UIFont 'Segoe UI' 8
    $cbLang.DrawMode = 'OwnerDrawFixed'
    $cbLang.IntegralHeight = $false
    $cbLang.Name = "cbLang"
    
    $cbLang.Items.AddRange(@("EN", "DE", "PT", "PHL"))
    $cbLang.Text = $script:GuideState.Language
    
    $cbLang.Add_SelectedIndexChanged({
        if ($this.Text -ne $script:GuideState.Language) {
            $script:GuideState.Language = $this.Text
            Show-GuideBubble
        }
    })
    $form.Controls.Add($cbLang)

    $lblText = New-Object System.Windows.Forms.Label
    $lblText.Location = New-Object System.Drawing.Point(10, 40)
    $lblText.Size = New-Object System.Drawing.Size(280, 75)
    $lblText.Font = Get-UIFont 'Segoe UI' 9
    $lblText.ForeColor = Get-Theme 'SecondaryText'
    $lblText.Name = "lblText"
    $form.Controls.Add($lblText)

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "Prev"
    $btnPrev.Size = New-Object System.Drawing.Size(60, 25)
    $btnPrev.Location = New-Object System.Drawing.Point(10, 125)
    $btnPrev.FlatStyle = 'Flat'
    $btnPrev.ForeColor = Get-Theme 'WhiteText'
    $btnPrev.BackColor = Get-Theme 'ButtonBg'
    $btnPrev.Name = "btnPrev"
    $btnPrev.Add_Click({ 
        $script:GuideState.CurrentStep--
        Show-GuideBubble 
    })
    $form.Controls.Add($btnPrev)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Next"
    $btnNext.Size = New-Object System.Drawing.Size(60, 25)
    $btnNext.Location = New-Object System.Drawing.Point(150, 125)
    $btnNext.FlatStyle = 'Flat'
    $btnNext.ForeColor = Get-Theme 'WhiteText'
    $btnNext.BackColor = Get-Theme 'AccentBlue'
    $btnNext.Name = "btnNext"
    $btnNext.Add_Click({ 
        $script:GuideState.CurrentStep++
        Show-GuideBubble 
    })
    $form.Controls.Add($btnNext)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Size = New-Object System.Drawing.Size(60, 25)
    $btnClose.Location = New-Object System.Drawing.Point(230, 125)
    $btnClose.FlatStyle = 'Flat'
    $btnClose.ForeColor = Get-Theme 'WhiteText'
    $btnClose.BackColor = Get-Theme 'ErrorColorAlt'
    $btnClose.Name = "btnClose"
    $btnClose.Add_Click({ Close-Guide })
    $form.Controls.Add($btnClose)
    
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = "1/10"
    $lblProgress.AutoSize = $true
    $lblProgress.Location = New-Object System.Drawing.Point(85, 130)
    $lblProgress.ForeColor = Get-Theme 'SecondaryText'
    $lblProgress.Font = Get-UIFont 'Segoe UI' 8
    $lblProgress.Name = "lblProgress"
    $form.Controls.Add($lblProgress)

    $form.AcceptButton = $btnNext
    $form.KeyPreview = $true
    $form.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq 'Right') {
            $script:GuideState.CurrentStep++
            Show-GuideBubble
            $e.Handled = $true
        }
        elseif ($e.KeyCode -eq 'Left') {
            if ($script:GuideState.CurrentStep -gt 0) {
                $script:GuideState.CurrentStep--
                Show-GuideBubble
            }
            $e.Handled = $true
        }
    })

    $dragHandler = { 
        param($s, $e) 
        if ($e.Button -eq 'Left') { 
            [Custom.Native]::ReleaseCapture()
            $f = if ($s -is [System.Windows.Forms.Form]) { $s } else { $s.FindForm() }
            if ($f) { [Custom.Native]::SendMessage($f.Handle, 0xA1, 0x2, 0) }
        } 
    }
    $form.Add_MouseDown($dragHandler)
    $lblTitle.Add_MouseDown($dragHandler)
    $lblText.Add_MouseDown($dragHandler)

    return $form
}

function Update-BubbleContent {
    param($Form, $Step, $StepIndex, $TotalSteps)

    $lang = $script:GuideState.Language
    $btnText = $script:ButtonLabels[$lang]

    $Form.Controls['lblTitle'].Text = $Step.Title
    $Form.Controls['lblText'].Text = $Step.Text
    $Form.Controls['lblProgress'].Text = "$($StepIndex + 1) / $TotalSteps"
    
    $cb = $Form.Controls['cbLang']
    if ($cb.Text -ne $lang) {
        $cb.Text = $lang
    }

    $btnPrev = $Form.Controls['btnPrev']
    $btnNext = $Form.Controls['btnNext']
    $btnClose = $Form.Controls['btnClose']

    $btnPrev.Text = $btnText.Prev
    $btnClose.Text = $btnText.Close

    $btnPrev.Enabled = ($StepIndex -gt 0)
    if ($StepIndex -eq $TotalSteps - 1) {
        $btnNext.Text = $btnText.Finish
        $btnNext.BackColor = Get-Theme 'SuccessColorAlt'
    } else {
        $btnNext.Text = $btnText.Next
        $btnNext.BackColor = Get-Theme 'AccentBlue'
    }
}

function PositionBubbleForm {
    param($Form, $Target)

    if (-not $Target) { return }

    $targetRect = $Target.RectangleToScreen($Target.ClientRectangle)
    
    $x = $targetRect.Right + 10
    $y = $targetRect.Top
    
    $screen = [System.Windows.Forms.Screen]::FromControl($Target)
    $bounds = $screen.WorkingArea

    if (($x + $Form.Width) -gt $bounds.Right) {
        $x = $targetRect.Left - $Form.Width - 10
    }

    if (($y + $Form.Height) -gt $bounds.Bottom) {
        $y = $bounds.Bottom - $Form.Height - 10
    }

    if ($x -lt $bounds.Left) { $x = $bounds.Left + 10 }
    if ($y -lt $bounds.Top) { $y = $bounds.Top + 10 }

    $Form.Location = New-Object System.Drawing.Point($x, $y)
}

#endregion