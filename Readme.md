# Reboot and Shutdown

This is a small .bat script written for my own needs.  
It is very simple. It reboots the computer and then shuts down.

This script does 5 things:  
- Writes a single byte to the registry, to know that it has just ran.  
- Restarts the computer.  
- Wait for Windows to start up and for the user to log in.  
- Deletes the byte saved in registry. It cleans up after itself.  
- Shut down the computer.  
  
It's that simple.  
This file can be ran from any location.  

## Copyright
Written by Ubehage 2026.  

## License
- MIT License. All code is free to use and modify.
