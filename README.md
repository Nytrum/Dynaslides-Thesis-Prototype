# Dynaslides-Thesis-Prototype

This repository represents all the code used and developed for the dynaslides thesis project.

## Project Structure
The structure is divided into four main directories:

### Backend
The backend is responsible for translating the powerpoint file and send it to the frontend. It can be used through a command-line interface with the need to deploy the frontend application.
The most important files and directories are:
 + *config* - This is where properties of the dynaslides application can be changed;
 + *js* - This is where all Javascript files, including the plugins, are located;
 + *app.js* - This is where the backend main implementation is done; It uses express.js to create an endpoint that provides the translation to the frontend implementation;
 + *dynaslides.js* - This is where all the logic for translating a presentation is located; 

### ConversionTests
The conversion tests are a way to see if the elements of the original presentation are equal or similar to the elements in the new presentation format.
### Frontend
The frontend is responsible for providing a user interface capable of sending and downloading requests made to the backend implementation. 
The most important files and directories are:
 + *src* - This is where all the implementation is present. The frontend takes advantage of the React Javascript Framework to display the user interface and send the request to the backend;
 + *src/App.js* - This is where the magic happens. All the functions to communicate with the backend application are here. The front-end application has the "http://localhost:7000/getDynaslidesPresentation" endpoint, and this must be changed to the endpoint of the backend to work.;

### Server
The server is an express.js implementation that sets up a socket.io server that will communicate between the master presentation and the client presentation. It has to be place in a public endpoint and this endpoint must be specified in the config file of the dynaslides application. 
## Installation and usage

### Command line version
The node client version is easy to use. To do it, it currently uses version 12.18.3 of node.js and uses the following commands:

> Clone repo from github.
>
> Unzip, go the Backend direcotry and use npm i
>
> Execute program with node dynaslides.js \<pptx location\> \<csv location\> ...

To enable the audience interaction feedback, please add to the previous command the tag “ai_on”.

The result presentation will go the directory **Backend/resultPresentations**
It needs to have one pptx file.

### Website version:
To set up the website environment locally, the presenter first has to do the following:

 + Launch a command-line, go to the Backend  directory and use npm i;
 + Execute the command node start;
 + Launch a different command-line, go to the Frontend directory and use npm i;
 + Execute the command node start;
 + Launch a different command-line, go to the Server directory and use npm i;
 + Execute the command node start;

If the presenter wants to have the application in a public endpoint, it has to change the endpoints in the frontend and backend application accordingly.

To use the application:

+ Use the “choose file” button to get the .pptx file.

Use the additional settings button/text to choose the different available options:

+ A button for the audience interaction feedback that will generate two folders (client and master) for the presentation, i.e, the one controlled by the presenter and the one used by the audience; 

+ A button to add the documents that have the data regarding the live visualization.

+ Click the download button, UNZIP the result folder and go to the index.html in the respective folder.

## Translation
The translation is currently an improved version of https://github.com/g21589/PPTX2HTML
## Plugins

To use the plugins, the presenter must first change the powerpoint presentation slide accordingly:

 + Create a text box, preferably a large one;
 + Right-click on the text box and choose edit **alt text**;

After that, the user can use the following available plugins:

 + **Live Website** - Insert %%liveviz("\<URL\>")%% on the **alt text** description of the text box.

 + **Live coding** - Insert %%livecode%% on the **alt text** description of the text box. This plugin does not have any additional settings.

 + **Live visualization** - Insert  %%liveviz("\<Filename\>.\<extension\>")%% on the **alt text** description of the text box, where \<Filename\> is the name of the data file and \<extension\> is the format of that said file. For example, "file.csv". Before translating the file, the presenter also must insert the respective data files on the website version's additional settings or give the additional option in the command-line version. Otherwise, the plugin will not work.

+ **Live Audience Interaction** - Insert  %%livequestion("\<Question\>")%% on the **alt text** description of the text box, where \<Question\> is the yes or no question the presenter wants to ask the audience. Before translating the file, the presenter also must activate the audience feedback on the website version's additional settings or give the additional option "ai_on" in the command-line version. Otherwise, the plugin will not work.
