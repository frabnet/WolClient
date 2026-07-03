# WolClient

This PowerShell script simplifies powering on remote computers over an OpenVPN/pfSense VPN.

Users simply run the script, follow the instructions, and wait for Remote Desktop to become available.

<img width="800" height="411" alt="WolClient demo" src="https://github.com/user-attachments/assets/2434e3bc-c68b-4671-80f1-904963986497" />

**Make users happy and save the planet. 🌸**
No more leaving office PCs running 24/7 just in case someone needs remote access.

Now also [available for Linux](https://github.com/frabnet/WolClient-Linux)! 🐧

## pfSense setup / server side

- Setup OpenVPN Server as you like.
  Hint: providing a DNS Server able to resolve dhcp clients hostnames can simplify things.
- Setup the firewall to allow OpenVPN clients reach pfSense (https) and the remote computer (rdp).
- Create a Wake On Lan user, without a certificate (it's not used in OpenVPN).
- Edit the Wake On Lan user, set "WebCfg - Services: Wake-on-LAN" under "Effective Privileges".

## WolClient setup / client side

- Install [OpenVPN](https://openvpn.net/community-downloads/) and copy the configuration file in %userprofile%\OpenVPN\config
- Download this repository and extract it to a known location.
- Create or edit `WolClientConfig.xml` and configure:
  - pfSense Hostname
  - pfSense Wake On Lan user
  - pfSense Wake On Lan password
  - Remote PC hostname/ip
  - Remote PC mac address
  - Remote PC username (used in RDP)
- Create a shortcut to `WolClient.cmd` for the user.
