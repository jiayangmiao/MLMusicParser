
# MLMusicParser
A piece of Java script that parses the information of all the music from game Idolm@ster Million Live! along with the official samples m4a links to them that are provided in provided in its official mobage.

It will also parse the information of the CDs. 

The code reads a certain input file (input.txt) and the user is responsible to feed to file with certain source code from the mobage's page that corresponds to a whole CD's information.
The decision was made since the in game audioroom provides all the information and it is not a static page. 

To use the app, rename the MLMusicInfoTemplate.xlsx into MLMusicInfo.xlsx and this is the file the script will write its outputs to.

To find the input, go to the mobage page of [audio room](http://imas.gree-apps.net/app/index.php/audio_room),
Use inspect elements or any other tools to get the source code and only copy the div with id tab-slide-area 
    `<div id="tab-slide-area">`
    
It resides in
    `<body><div id="gree-app-container"><div id="wrapper"><div class="main-bg">`
    
Copy the element from chrome will help you copy the whole div easily. Then paste it into input.txt.
