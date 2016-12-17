# imhome
I always told myself that I should find a way to automatically launch some applications that require bandwidth from my computer when I left home.
The purpose of this software should exactly be that one, when you left home, and some device comes with you (ex. your smartphone)
the software recognizes that missing, and launches some application.

Well than, now is time for some tecnical description:

With 3 threads the software scan your whole LAN and using arp it looks for the MAC address of the connected devices, if it founds the one you have inserted it will close the application(s) it has launched before.
When the selected MAC address disappear from the network it will launch the selected application(s)
Each 50 seconds a timer restarts the threads to keep the scanning.  



Right now you have no choise than launch at first mipony, then if no downloads start, imhome will replace it with utorrent.
Of course as soon as you are back home it closes the running app, so that your bandwidth is free for you.

TO DO:

Use different timer inteval when the device is connected to the network (a long one) and when the device is disconnected from the network (short one);

Add the possibility to change that intevals from a GUI;

Add the possibility to change the software you want to launch (righ now first it launches mipony, then if no downloads start, it replaces it with utorrent) and the ability to insert multiple MAC devices.
