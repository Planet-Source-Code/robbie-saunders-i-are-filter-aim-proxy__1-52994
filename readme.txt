=================================
 How to Setup
=================================
   - Open AIM Preferences
	- Click "Sign On/Off" (on the left)
	- Click "Connection" (bottom right)
		- Check "Connect using proxy"
			- Choose "SOCKS 4"
			- Host = "localhost"
			- Port = "3333" (or whatever you use)
	- Connect as normal

=================================
 How to use Scripts
=================================
   - Scripts must be saved in the scripts folder
   - You must also enable scripts in the automation menu
   - From then on it's just like using a normal command ie.
	aim.script.multiwarn
     would execute a script called `multiwarn` if it exists


=================================
 How to write Scripts
=================================
   - For each line, put what text you want sent in an im ie.
	hey gringo i'm gonna warn you
	aim.warn.user
     would send someone the message "hey gringo i'm gonna warn you"
     then it would warn them hah
   - Save whatever extension you want, it only looks at everything
     to the left of the first period