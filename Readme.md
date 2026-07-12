# Reboot and Shutdown (1.1.0)

This is a small program, written for my own needs.  
It comes in 2 flavors: A batch script, and a win32 executable written in VB6.  
Both does roughly the same:  
They are very simple. They reboot the computer and then shuts down.  

The batch script does 5 things:  
- Writes a single byte to the registry, to know that it has just ran.  
- Restarts the computer.  
- Wait for Windows to start up and for the user to log in.  
- Deletes the byte saved in registry. It cleans up after itself.  
- Shut down the computer.  
  
The executable does not save anything to the registry, but otherwise performs the same.  
The executable will also attempt to install updates during reboot on newer Windows versions.  

It's that simple.  
The files can be run from any location. 100% portable.  

## Recent changes and fixes
### Changes to the application:
- 1.1.0
  - Added a short delay of 5 seconds before shutting down after reboot, with a visual countdown.
- 1.0.1
  - Program would shut down instead of rebooting on Windows 10/11.
  
## Copyright
Written by Ubehage 2026.  

## License
- MIT License. All code is free to use and modify.
