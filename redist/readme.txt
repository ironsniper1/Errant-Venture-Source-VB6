Use Sample.exe to run the XvT Client






----------------------------------------
Configuring
----------------------------------------

Sample.exe should work right out of the box, but depending on the server you want to connect to you may need to do a few things first


-OverRide function - if you are connecting from inside a Lan and you internal IP address always shows up in EV, you can insert your external IP Address or a redirector from no-ip into or.dat to correct this (you can find your external IP Address by going to www.whatismyip.com

-multiple hosts list, you can change the order or add/remove addresses by editing the hosts.dat file

-default port override... if for some reason 2020 is no longer acceptable, the server host can edit the first line in remport.dat to change the port the server listens to for incomming connections


- if a server host has changed the port that the server listens to and hasn't mapped it to external port 2020, then users can edit the first line in ports.dat to change the port the client tries to connect to


----------------------------------------
loging in
----------------------------------------

the clients connect to the server on port 2020.

to login press connect - the default address 
is the central server, 

you will then see 2 boxes,
userid and password, 

if you have not registered in the Errant Venture System yet, you will need to

to do this enter your uid, and pwd you desire, then
press login/register

you will see an errormessage - press OK to go to the register 
window, then re enter your username then password

press ok... you will either recieve a message that you have succesfully registered,
or your id is already registered, in which case choose another.


----------------------------------------
Using the client
----------------------------------------

Players list
----------------------------------------

the box on the left side of the screen will show a list of players currently online

Preceeding the name will be a single letter in brackets

L=Lurking
P=Playing
A=Away
H=Hosting a room
G=Guest in a room

you can set your status to away by pressing the away button below the players box

to page a user (private IM Message) either dbl click a players name, or select
the player and press page.

when you have a player selected the tool text tip for the whole box will display the ip
address for that player


----------------------------------------
TROUBLE SHOOTING
----------------------------------------

1) make sure you have copied all of the .ocx files to which ever of the following exists 

C:\windows\system32

or

C:\winnt\system32


2) then you may have to register the AResize.ocx... you can do this by using the run menu item from you start menu and typin in:


regsvr32 C:\WINDOWS\system32\AResize.ocx

or 

regsvr32 C:\WINNT\system32\AResize.ocx

which ever exists

3) try a reboot to reload all the system files

read the errorlog file and see if there is some clue why, if not post it to the NRSD Errant Venture commlink or the battlestats msg board and ask for assistance














Game Room List
----------------------------------------

the box on the right holds a list of avalible games

to see who is in a game, click on the game, the tool text tip
of the box will now show a list of players

to join a game, select the game you wish to join, the join
button will become avalible

to host a game, just press the host button, the configure game button will come up, and  
you can select the number of player, the title of the game room, and the game to be played

on pressing ok, the gameroom will be spawned and other players will be able to see it


----------------------------------------
mute
----------------------------------------
there is a sound played when a user enters and leaves the lobby, and another set for 
the game rooms, to disable these, check the mute button below the game room list box

-----------------------------------------
gamerooms
-----------------------------------------

gamerooms are hosted on the hosting players machine
in order for other members to connect to the hosting 
players game room, ports 1001, 2403-2404 must be allowed
to accept incomming connections


If you cannot connect to the chat portion of the game room (players window does not get names of players in gameroom)
page the host and tell them that you are in the game room and to page you just before they launch, so that you can click IP In which will manually launch the game on your end, which will allow you to connect to the game without properly connecting to the gameroom.


-----------------------------------------

SampleEcho.exe is a server, I have this running on my server, 
and that is what sample.exe connects too

you can use this to set up you own local network if you wish, 
however the default address in the clients will always point to mine,
if you want to run you own, you will have to let your friends know
what your address is. 

this might be useful for University wide lan parties and the like

Port 2020 must be able to accept incomming connections.


--------------------------------------------
Using a router with Errant Venture
--------------------------------------------

A router is a device that sits between your DSL or Cable Modem and your computers and pretends to be a computer to the modem, and a modem to your computers

The consequence of this is that though the router knows which computer sends a request for a connection outside it's network and will know who to send the replys that come back from the internet to, requests for new connections from the internet will be addressed to the router address, not the address of the computers and thus it ignores them.

You can do 2 things to tell the router where connections are supposed to go, Port Forwarding and DMZ

Port Forwarding tells the router to send traffic on certain ports to a designated computer. ie ports 2300 to 2400 - TCP - 192.168.1.11

DMZ simply copies all the traffic that comes in, and sends it all to the designated computer.


--------------------------------------------
setting up a router - port forwarding
--------------------------------------------

1 - on the computer you wish to set up, goto command prompt and run ipconfig /all

2 - copy down DNS Servers

3 - set the network adapter on this computer to a static IP in the range of like 192.168.1.11(if other is not set to this) 

4 - use subnet mask 255.255.255.0, gateway 12.168.1.1, and use the DNS servers you coppied down

5 - reboot

6 - go into router settings - advanced - port forwarding

7 - for ipaddress 192.168.1.11 forward ports 1001, 2300-2400, 47624 and 6073. once for TCP once for UDP

--------------------------------------------
setting up a router - DMZ
--------------------------------------------

1 - on the computer you wish to set up, goto command prompt and run ipconfig /all

2 - copy down DNS Servers

3 - set the network adapter on this computer to a static IP in the range of like 192.168.1.10(if other is not set to this) 

4 - use subnet mask 255.255.255.0, gateway 12.168.1.1, and use the DNS servers you coppied down

5 - reboot

6 - go into router settings - advanced - DMZ

7 - Set the DMZ to goto the ip 192.168.1.10 and Enable DMZ

