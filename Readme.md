# Reboot and Shutdown

This is a small program, written for my own needs.  
It comes in 2 flavors: A batch script, and a win32 executable written in VB6.  
Both does exactly the same:  
They are very simple. They reboot the computer and then shuts down.  

The batch script does 5 things:  
- Writes a single byte to the registry, to know that it has just ran.  
- Restarts the computer.  
- Wait for Windows to start up and for the user to log in.  
- Deletes the byte saved in registry. It cleans up after itself.  
- Shut down the computer.  
  
The executable does not save anything to registry, but otherwise functions exactly the same.  

It's that simple.  
The files can be run from any location. 100% portable.  

## Copyright
Written by Ubehage 2026.  

## License
- MIT License. All code is free to use and modify.
