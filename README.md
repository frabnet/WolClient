# WolClient

A script to simplify power on of remote computers while connected via OpenVPN/pfSense VPN. Make users happy and save the planet :blossom:

User just need to run the script, follow instructions and wait for RDP to come  up.

## pfSense setup

- Setup OpenVPN Server as you like.

  Hint: providing a DNS Server able to resolve dhcp clients hostnames can sometimes simplify things.

- Setup the firewall to allow OpenVPN clients reach pfSense (https) and the remote computer (rdp).

- Create a Wake On Lan user, without a certificate (it's not used in OpenVPN).

- Edit the Wake On Lan user, set "WebCfg - Services: Wake-on-LAN" under "Effective Privileges".

## WolClient setup

- Download this repository and save in a known directory.
- Create/edit WolClientConfig.xml with required settings:
  - pfSense Hostname
  - pfSense Wake On Lan user
  - pfSense Wake On Lan password
  - Remote PC hostname/ip
  - Remote PC mac address
  - Remote PC username (used in RDP)
- Make a link to the WolClient.cmd file for the user.

