
<p align="center">
  <a href="" rel="noopener">
 <img width=400px height=500px src="https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/interface-start.png" alt="Entropia Dashboard"></a>
</p>

<h3 align="center">Entropia Dashboard</h3>

<div align="center">

  [![Homepage](https://img.shields.io/badge/Homepage-gray)](https://immortal-divine.github.io/Entropia_Dashboard/) 
  [![Download](https://img.shields.io/badge/Download-Now-blue)](https://github.com/Immortal-Divine/Entropia_Dashboard/raw/refs/heads/main/Entropia%20Dashboard.exe) 
  [![Donate](https://shields.io/badge/ko--fi-donate-ff5f5f?logo=ko-fi)](https://ko-fi.com/U7U61S0EGT) 

</div>

---
A powerful, free, and open-source application designed to streamline your Flyff gameplay.

Written in PowerShell, it offers robust client management, automated launch and logins, and a customizable F-Tool/Macro to enhance your efficiency without compromising security. 
And many more!

- [About](#about)
- [Getting Started](#getting_started)
- [Usage](#usage)
- [Built Using](#built_using)

## About <a name = "about"></a>

> Interface

![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/custom-start.png) ![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/interface.png)

> Settings

![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/settings-general.png)
![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/settings-login.png)
![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/custom-profile.png)

> Ftool

![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/ftool.png)

> Disconnect System

![](https://raw.githubusercontent.com/Immortal-Divine/Entropia_Dashboard/refs/heads/main/Preview/reconnect.png)
 


  One-Click Login & Launch
  Forget typing passwords 8 times. Setup your team profiles once, and launch/login your entire party with a single click.
   * Smart Patching: Detects when patching is done and launches immediately.
   * Auto-Entry: Handles Server, Channel, Character selection!

  Smart Reconnect System
  Internet hiccup? AFK disconnect? We got you.
   * The dashboard monitors your connection in real-time. If a client drops, it automatically handles the "Disconnect" dialog and logs you back in.

  F-Tool & Macro
  Built-in tools that work in the background.
   * Advanced F-Tool: Create unlimited key spammers. They attach to specific windows so you can alt-tab freely without stopping your rotation.
   * Sequence Macros: Need complex combos? Build timed sequences with holds, waits, and loops.
   * Overlay Mode: Draggable mini-windows that snap to your game client.

  Chat Commander
  Type commands like /PartyInvite or shout trade messages across multiple clients with global hotkeys. Supports variables for dynamic messaging!

  Built-in Knowledge Base
   * Online/Offline Wiki: Access guides instantly.

## Getting Started <a name = "getting_started"></a>

<<<<<<< HEAD
## Getting Started <a name = "getting_started"></a>

=======
>>>>>>> 8016ac75b3965a6d829b20c559c99d30d623b50b

+ Process list

	+ Displays the current running clients separated by profiles.

	+ Order can be changed with the Remove Key.

	+ Indicator for new message ingame or if client got disconnected.

	+ Multi selection possible.

	+ Right click for context menu with additional features.

	+ Lists the current state of the window (e.G. Loading, Minimized, Normal).

	+ Selecting one or more entrys is mandatory for other features.


+ Profiles

	+ You can create a junction/copy of your main folder (~300mb per Profile).

	+ Profiles can also be manually added.

	+ Every profile has it's own ingame settings.

	+ Every profile will be patched with a single patch (except for neuz.exe, must be patched in main folder).

	+ Select Profile to use as default profile.


+ Launch

	+ Can start the desired total amount of clients without loading errors.

	+ Can start one instance of a specific Profile with right-click

	+ Works in the background

	+ Saved One-Click Setup can be applied to start and login your final Setup.

	+ Clicking on the Launch button again aborts the launch operation.
   

+ Login

	+ The login takes the # of the process list to log the same # as position in the nickname list of the client. (Nickname list with 10 entrys mandatory)
	 + \# 4 is nickname # 4 etc.

	+ Once the login process is completed, it can click on 'Start' for collecting, can be enabled in the Settings. 

	+ Login only works for the first 10 clients of each profile.

	 + You can select which server/channel/character to login in the Settings.

	 + Mouse movement aborts the login operation.


+ Ftool

	+ Remembers the character and its settings.

	+ Gets pinned to the window.

	+ Up to 10 Keys can be set with a minimum delay of 10ms.
	 + Multiple Ftools are possible.

	+ Hotkey Manager for every individual Ftool Instance
	 + Master Hotkey Toggle blocks other Ftool Hotkeys
	 + Can also be toggled with a hotkey with the ⌨ icon.

	+ Custom position for each client

	+ Can bind ANY key combination to any Ftool or Hotkey

+ Terminate

	+ Close the selected clients instantly


## Usage <a name="usage"></a>

+ Download and start Entropia Dashboard.exe

+ Open Settings

	+ Select the Launcher.exe of your Flyff Server
	+ Save

+ Right-Click Launch and start 1 Client

+ To login you have to save your logins in your Client

	+ To use this feature your Server must support nickname lists
	+ If you want to login more than 5 different account on your Client, you must add exactly 10 logins to the list
	+ Login will happen automatically and use the login settings of the Dashboard if provided
	+ Default resolution is for 1024x768, if you want to use another size you must set the coordinates for First Nickname and Scroll Down Arrow

+ Select your Client in the Dashboard and click Ftool

	+ This Ftool is very powerful. Just imagine something, it probably can do it.

+ The One-Click Setup

	+ Start a set of your desired clients and log them in
	+ You can order around the clients in the dashboard by selecting an entry and pressing the delete key
	+ Save this list in the dashboard settings
	+ Now the dashboard will always start and login missing clients of your saved list compared to current running clients

+ Profiles

	+ A profile can be either added manually or by using the create button in the settings
	+ If created with the dashboard, your profile will share the base data with your main launcher folder, but still have its own dashboard AND INGAME settings
	+ To patch all your profiles, you just need to patch with the main launcher. Press Copy Neuz.exe after the patch to complete it for all profiles
	+ All features of the Dashboard are linked with profiles, you don't have to use them. But your expierence will be way better



+ Check the tooltips! All UI elements have them on mouse hover!
+ Open the guide for a more defauled explanation!

<<<<<<< HEAD

=======
>>>>>>> 8016ac75b3965a6d829b20c559c99d30d623b50b
## Built Using <a name = "built_using"></a>
- [Powershell 5.1]([https://www.mongodb.com/](https://github.com/PowerShell/PowerShell))


<a href='https://ko-fi.com/U7U61S0EGT' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://storage.ko-fi.com/cdn/kofi6.png?v=6' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>


# [Entropia Flyff](https://entropia.fun/)

# [Divine Discord](https://discord.gg/zbcVRsC9uN)


