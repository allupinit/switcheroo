# switcheroo
Switch (or add) an Exchange SMTP Alias temporarily.

The original version has been tested by me, and is fully operational with no known glitches.

Will prompt for e-mail address, check exchange to see who it belongs to.  
If it belongs to someone it will remove it from that person and add it to whomever you choose.  
If it does not currently exist on the server it will and add it to whomever you choose.  
The script will pause, when you press any key, it will undo the changes.  


The prompt for who to move/assign the address to can be either a username or e-mail address.
The script will query Exchange for an account matching your input, so either will do.
