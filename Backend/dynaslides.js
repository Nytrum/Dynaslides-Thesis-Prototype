// var arguments = processCommandLine(); // Setup the command line inpu
// startDynaslides(arguments); // Reads the powerpoint file and starts the slide processing
const DS_CONFIG = require('./config/dynaslides-config.json');
const path = require('path');
const CSV = require("csv-string");
const process = require("process");
const rdl = require("readline");
const fs = require("fs");
const Jimp = require("jimp");
const sharp = require("sharp");
const JSZip = require("jszip");
const { tXml } = require("./js/tXmlCLI.min.js");
const colz = require("./js/colz.class.min.js");
const tinycolor = require("./js/tinycolorNode.js");

//---------------------- PLUGINS ------------------------

const LiveCode = require("./js/plugins/livecode");
const LiveViz = require("./js/plugins/liveviz");
const LiveQuestion = require("./js/plugins/livequestion");

//---------------------- VARIABLES ------------------------
var size = 100;
var cursor = 3;
var slideSize = null;
var currentProgress = 0;
var current_slide = 0;
var data_files_names = [];
var data_files_data = [];
var mainthemeContent = null;
var currentthemeContent = null;
var log_info = "";
var chartID = 0;
var powerPointFile = null;
var audience_interaction_plugin = false;
var plugins = [];

var biggestMasterLayoutOrder = 0;
var styleTable = {};

if (require.main === module) {

  var myArgs = processCommandLine();

  if(powerPointFile){
    fs.readFile(powerPointFile, function (error, data) {
      for (let i = 0; i < myArgs.length; i++) {
        data_files_names.push(myArgs[i]);
        data_files_data.push(
          CSVArrayToJSON(
            CSV.parse(
              fs.readFileSync(myArgs[i], { encoding: "utf8", flag: "r" })
            )
          )
        );
      }
      processPPTX(data).then(function (slides) {
        if(audience_interaction_plugin){
          var resulting_zip = getCompressedFileWithAudienceInteractionCL(slides,myArgs);
        } else{
          var resulting_zip = getCompressedFileCL(slides,myArgs);
        }
        resulting_zip
          .generateNodeStream({ type: "nodebuffer", streamFiles: true })
          .pipe(fs.createWriteStream("resultPresentations/"+ powerPointFile.split('.').slice(0, -1).join('.') + '.zip'))
          .on("finish", function () {
            console.log(powerPointFile.split('.').slice(0, -1).join('.') + '.zip' + " was saved in the resultPresentation directory.");
            if (log_info != "") {
              fs.writeFile("resultPresentations/"+powerPointFile.split('.').slice(0, -1).join('.')+"_errors.txt", log_info, function (err) {
                if (err) return console.log(err);
                console.log(
                  "There were errors in the build. " + powerPointFile.split('.').slice(0, -1).join('.')+"_errors.txt" +" was created in resultPresentations directory!"
                );
              });
            }
        });
      });
    });
  }
  
}

/**
 * Holds JavaScript settings and other informations for Dynaslides application with the website version.
 * @param {file} pptxfile File info and buffer of the original powerpoint presentation
 * @param {array} args Custom options of the additional settings
 * @namespace
*/
class Dynaslides {
  constructor(pptxfile,args) {
    this.powerPointFile = pptxfile;
    this.args = args;
  }
  async startDynaslides() {
    for (let i = 0; i < this.args.data_files.length; i++) {
      data_files_names.push(this.args.data_files[i].originalname);
      data_files_data.push(
        CSVArrayToJSON(
          CSV.parse(
            this.args.data_files[i].buffer.toString()
          )
        )
      );
    }

    return processPPTX(this.powerPointFile);
  }

    /**
   * This is where the file format is assembled. All the folder structure is done here
   * @param {text} slides All the html constructed from the translation
   * @return
   *   Returns the zip file object with the presentation.
   */
  getCompressedFile(slides) {
    var cssText = fs.readFileSync("css/pptx2html.css");
    var headHtml = "<style>" + cssText + "</style>";
    var bodyHtml = "<div id='slides' class='slides'>" + slides + "</div>";
    var html = createHTMLrevealFile(headHtml, bodyHtml);
    var zip = getZipOfFolder("reveal");
    zip.file("index.html", html);
    var data = zip.folder("data");
    for (let i = 0; i < this.args.data_files.length; i++) {
      data.file(this.args.data_files[i].originalname, this.args.data_files[i].buffer.toString());
    }
    var js_zip = zip.folder("js");
    var socketio_data = fs.readFileSync("./js/socket.io.js");
    js_zip.file("socket.io.js", socketio_data);
    var images_zip = zip.folder("images");
    var favicon_data = fs.readFileSync("./public/images/dynaslideslogo.ico");
    images_zip.file("favicon.ico", favicon_data);
    return zip;
  }

  /**
   *
   * This is where the file format with the audience feedback is assembled. All the folder structure is done here.
   *
   * @param {text} slides All the html constructed from the translation.
   * @return
   *   Returns the zip file object with the presentation.
   */

  getCompressedFileWithAudienceInteraction(slides) {
    var cssText = fs.readFileSync("css/pptx2html.css");
    var headHtml = "<style>" + cssText + "</style>";
    var bodyHtmlMaster = masterAIPanel() + "<div id='slides' class='slides'>" + slides + "</div>";
    var bodyHtmlClient = clientAIPanel() + "<div id='slides' class='slides'>" + slides + "</div>";
    var html_master = createHTMLrevealFile(headHtml, bodyHtmlMaster, "master");
    var html_client = createHTMLrevealFile(headHtml, bodyHtmlClient, "client");
    var zip = new JSZip();
    var master_folder = insertFolderInZip(zip, "reveal", "master");
    var client_folder = insertFolderInZip(zip, "reveal", "client");
    master_folder.file("index.html", html_master);
    client_folder.file("index.html", html_client);
    var data_master = master_folder.folder("data");
    var data_client = client_folder.folder("data");
    for (let i = 0; i < this.args.data_files.length; i++) {
      data_master.file(this.args.data_files[i].originalname, this.args.data_files[i].buffer.toString());
      data_client.file(this.args.data_files[i].originalname, this.args.data_files[i].buffer.toString());
    }
    var js_zip_master = master_folder.folder("js");
    var js_zip_client = client_folder.folder("js");
    var socketio_data = fs.readFileSync("./js/socket.io.js");
    js_zip_master.file("socket.io.js", socketio_data);
    js_zip_client.file("socket.io.js", socketio_data);
    var master_images_zip = master_folder.folder("images");
    var client_images_zip = client_folder.folder("images");
    var favicon_data = fs.readFileSync("./public/images/dynaslideslogo.ico");
    master_images_zip.file("favicon.ico", favicon_data);
    client_images_zip.file("favicon.ico", favicon_data);
    return zip;
  }
}

    /**
   *
   * This is where the file format is assembled. All the folder structure is done here
   *
   * @param {text} slides All the html constructed from the translation
   * @param {array} myArgs Custom arguments of the command line. Data files and options
   * @return
   *   Returns the zip file object with the presentation from the command line version.
   */

function getCompressedFileCL(slides,myArgs) {
  var cssText = fs.readFileSync("css/pptx2html.css");
  var headHtml = "<style>" + cssText + "</style>";
  var bodyHtml = "<div id='slides' class='slides'>" + slides + "</div>";
  var html = createHTMLrevealFile(headHtml, bodyHtml);
  var zip = getZipOfFolder("reveal");
  zip.file("index.html", html);
  var data = zip.folder("data");
  for (let i = 1; i < myArgs.length; i++) {
    data.file(myArgs[i], fs.readFileSync(myArgs[i]));
  }
  var js_zip = zip.folder("js");
  var socketio_data = fs.readFileSync("./js/socket.io.js");
  js_zip.file("socket.io.js", socketio_data);
  var images_zip = zip.folder("images");
  var favicon_data = fs.readFileSync("./public/images/dynaslideslogo.ico");
  images_zip.file("favicon.ico", favicon_data);
  return zip;
}

    /**
   *
   * This is where the file format is assembled with audience feedback activated. All the folder structure is done here
   *
   * @param {text} slides All the html constructed from the translation
   * @param {array} myArgs Custom arguments of the command line. Data files and options
   * @return
   *   Returns the zip file object with the presentation from the command line version.
   */

function getCompressedFileWithAudienceInteractionCL(slides,myArgs) {
  var cssText = fs.readFileSync("css/pptx2html.css");
  var headHtml = "<style>" + cssText + "</style>";
  var bodyHtmlMaster = masterAIPanel() + "<div id='slides' class='slides'>" + slides  +"</div>";
  var bodyHtmlClient = clientAIPanel() + "<div id='slides' class='slides'>" + slides + "</div>";
  var html_master = createHTMLrevealFile(headHtml, bodyHtmlMaster,"master");
  var html_client = createHTMLrevealFile(headHtml, bodyHtmlClient,"client");
  var zip = new JSZip();
  var master_folder = insertFolderInZip(zip,"reveal","master");
  var client_folder = insertFolderInZip(zip,"reveal","client");
  master_folder.file("index.html", html_master);
  client_folder.file("index.html", html_client);
  var data_master = master_folder.folder("data");
  var data_client = client_folder.folder("data");
  for (let i = 1; i < myArgs.length; i++) {
    data_master.file(myArgs[i], fs.readFileSync(myArgs[i]));
    data_client.file(myArgs[i], fs.readFileSync(myArgs[i]));
  }
  var js_zip_master = master_folder.folder("js");
  var js_zip_client = client_folder.folder("js");
  var socketio_data = fs.readFileSync("./js/socket.io.js");
  js_zip_master.file("socket.io.js", socketio_data);
  js_zip_client.file("socket.io.js", socketio_data);
  var master_images_zip = master_folder.folder("images");
  var client_images_zip = client_folder.folder("images");
  var favicon_data = fs.readFileSync("./public/images/dynaslideslogo.ico");
  master_images_zip.file("favicon.ico", favicon_data);
  client_images_zip.file("favicon.ico", favicon_data);
  return zip;
}

  /**
 *
 * This is where the main html file is assembled. All the new slide information is done here
 *
 * @param {text} css css of the translation
 * @param {text} html HTML from the translation
 * @param {boolean} ai_option Says if the audience feedback is activated
 * @return
 *   Returns the html file object with the presentation.
 */

function createHTMLrevealFile(css, html,ai_option=null) {
  var audience_interaction_options = '';
  var scriting = ''
  if(ai_option){
    if(ai_option == "master"){
      audience_interaction_options = `
      multiplex: {
        secret: '${DS_CONFIG.audience_interaction_master_io_secret_key}', // Obtained from the socket.io server. Gives this (the master) control of the presentation
        secret_c: null,
        id: '${DS_CONFIG.audience_interaction_master_io_id}', // Obtained from socket.io server
        id_c: '${DS_CONFIG.audience_interaction_client_io_id}', // Obtained from socket.io server
        url: '${DS_CONFIG.audience_interaction_io_url}' // Location of socket.io server
      },`;
      scriting = masterJSLogic();
    }else if(ai_option=="client"){
      audience_interaction_options = `
      multiplex: {
        secret: null, 
        secret_c: '${DS_CONFIG.audience_interaction_client_io_secret_key}', 
        id: '${DS_CONFIG.audience_interaction_master_io_id}', // Obtained from socket.io server
        id_c: '${DS_CONFIG.audience_interaction_client_io_id}', // Obtained from socket.io server
        url: '${DS_CONFIG.audience_interaction_io_url}' // Location of socket.io server
      },`;
      scriting = clientJSLogic();
    }
  }

  return `<!doctype html>
  <html lang="en">
  
    <head>
      <meta charset="utf-8">
  
      <title>Slides</title>
  
      <meta name="apple-mobile-web-app-capable" content="yes">
      <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <link rel="icon" href="images/favicon.ico">
      <link rel="stylesheet" href="reveal/css/reset.css">
      <link rel="stylesheet" href="reveal/css/reveal.css">

      <link rel="stylesheet" href="reveal/lib/css/monokai.css">
      ${css}
    </head>
  
    <body>
  
      <div class="reveal">
  
        <!-- Any section element inside of this container is displayed as a slide -->
        ${html}
  
      </div>
  
      <script src="reveal/js/reveal.js"></script>
      <script src="js/socket.io.js"></script>
  
      <script>
  
        // More info https://github.com/hakimel/reveal.js#configuration
        Reveal.initialize({
          width: ${+slideSize.width},
          height: ${+slideSize.height},
          center: true,
          hash: true,
  
          transition: 'none', // none/fade/slide/convex/concave/zoom
          
          ${audience_interaction_options}
          
          // More info https://github.com/hakimel/reveal.js#dependencies
          dependencies: [	
            { src: 'reveal/plugin/highlight/highlight.js', async: true },
            { src: 'reveal/plugin/d3/d3.min.js', async: true },
            { src: 'js/socket.io.js', async: true },
            { src: 'reveal/plugin/multiplex/master.js', async: true },
            { src: 'reveal/plugin/multiplex/client.js', async: true }
          ]
        });
  
      </script>

      <script>
      var confirmDone = false;
      function runJavascriptCode(e){
        if (!confirmDone) {
          confirm("Are you sure you want to run the code displayed? Make sure you know what the output is!");
          confirmDone = true;
        }
        var code = e.parentElement.children[1].children[0].children[0].innerText;
        console.old = console.log;
        console.log = function () {
          var output = "", arg, i;
  
          for (i = 0; i < arguments.length; i++) {
            arg = arguments[i];
            output += "<tr><td class='log-" + (typeof arg) + "'>";
  
            if (
              typeof arg === "object" &&
              typeof JSON === "object" &&
              typeof JSON.stringify === "function"
            ) {
              output += JSON.stringify(arg);
            } else {
              output += arg;
            }
  
            output += "</td></tr>";
          }
  
          e.parentElement.children[3].children[0].children[0].innerHTML += output;
          console.old.apply(undefined, arguments);
        };
  
        eval(code);
        var i = document.createElement('iframe');
        i.style.display = 'none';
        document.body.appendChild(i);
        window.console = i.contentWindow.console;
      }
      ${scriting}
      </script>
    </body>
  </html>`;
}

function masterJSLogic(){
  return `

  (function() {
    window.presentation_type = "master";
    window.presentation_name = "${ powerPointFile.split('.').slice(0, -1).join('.') }";
    var multiplex = Reveal.getConfig().multiplex;
    var socketId = multiplex.id_c;
    var socket = io.connect(multiplex.url);
    var slidesAITracker = [];
    socket.on(multiplex.id_c, function(data) {
      // ignore data from sockets that aren't ours
      if (data.socketId !== socketId) { return; }
      if( window.location.host === 'localhost:1947' ) return;
      saveInformation(data);
  });

  function download(content, fileName, contentType) {
      var a = document.createElement("a");
      var file = new Blob([content], {type: contentType});
      a.href = URL.createObjectURL(file);
      a.download = fileName;
      a.click();
  }

  function saveInformation(data){
      var userSlideInfo = slidesAITracker.filter(e => e.id == data.data.slide);
      if(data.data.type == 3){
        var keys = Object.keys(window.received_data[0]);
        window.received_data[window.received_data.findIndex( ({ x }) => x === data.data.message )].y++;
        if (userSlideInfo.length == 0) {
          slidesAITracker.push({id:data.id,slides:[{slide:data.data.slide,likes:0,dislikes:0,messages:[],question:data.data.message}]});
          updateChart();
        }else{
          var slideInfo = userSlideInfo[0].slides.filter(e => e.slide == data.data.slide);
          if (slideInfo.length == 0) {
            userSlideInfo[0].slides.push({slide:data.data.slide,likes:0,dislikes:0,messages:[],question:data.data.message});
            updateChart();
          }else{
            if(slideInfo[0].question == ''){
              slideInfo[0].question = data.data.message;
              updateChart();
            }
          }
        }
      }else{
        if (userSlideInfo.length == 0) {
          if(data.data.type == 0){
              slidesAITracker.push({id:data.id,slides:[{slide:data.data.slide,likes:1,dislikes:0,messages:[],question:''}]});
          }else if(data.data.type == 1){
              slidesAITracker.push({id:data.id,slides:[{slide:data.data.slide,likes:0,dislikes:1,messages:[],question:''}]});
          }else if(data.data.type == 2){
              slidesAITracker.push({id:data.id,slides:[{slide:data.data.slide,likes:0,dislikes:0,messages:[data.data.message],question:''}]});
          }
        }else{
          var slideInfo = userSlideInfo[0].slides.filter(e => e.slide == data.data.slide);
          if (slideInfo.length == 0) {
            if(data.data.type == 0){
              userSlideInfo[0].slides.push({slide:data.data.slide,likes:1,dislikes:0,messages:[],question:''});
            }else if(data.data.type == 1){
              userSlideInfo[0].slides.push({slide:data.data.slide,likes:0,dislikes:1,messages:[],question:''});
            }else if(data.data.type == 2){
              userSlideInfo[0].slides.push({slide:data.data.slide,likes:0,dislikes:0,messages:[data.data.message],question:''});
            }
          }else{
            if(data.data.type == 0 && slideInfo[0].likes == 0){
              slideInfo[0].likes++;
            }else if(data.data.type == 1 && slideInfo[0].dislikes == 0){
              slideInfo[0].dislikes++;
            }else if(data.data.type == 2){
              slideInfo[0].messages.push(data.data.message);
            }
          }
        }        
      }
  }

  function updateChart(){
    var parentDiv = document.getElementById("liveviz");
    var c_width = parentDiv.clientWidth;
    var c_height = parentDiv.clientHeight * 0.8;
    // set the dimensions and margins of the graph
    var margin = {top: 20, right: 20, bottom: 30, left: 40},
    width = c_width - margin.left - margin.right,
    height = c_height - margin.top - margin.bottom;

    // set the ranges
    var x = d3.scaleBand()
                .range([0, width])
                .padding(0.1);
    var y = d3.scaleLinear()
                .range([height, 0]);
    d3.select("#liveviz_svg").remove();
    var svg = d3.select("#liveviz").append("svg")
        .attr("id","liveviz_svg")
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom + 30)
        .append("g")
        .attr("transform", 
                "translate(" + margin.left + "," + (margin.top + 30) + ")");
    
    x.domain(window.received_data.map(function(d) { return d.x; }));
    y.domain([0, d3.max(window.received_data, function(d) { return d.y; })]);

    var color = ["#087108", "#ff0000"];

    svg.selectAll(".bar")
        .data(window.received_data)
        .enter().append("rect")
        .attr("class", "bar")
        .style("fill", function(d, i) {
          return color[i];
        })       
        .attr("x", function(d) { return x(d.x); })
        .attr("width", x.bandwidth())
        .attr("y", function(d) { return y(d.y); })
        .attr("height", function(d) { return height - y(d.y); })           


    svg.selectAll("text.bar")
      .data(window.received_data)
      .enter().append("text")
      .attr("class", "bar")
      .style('transform', 'translate(0,-10px)')
      .attr("text-anchor", "middle")
      .attr("x", function(d) { return x(d.x) + x.bandwidth()/2; })
      .attr("y", function(d) { return y(d.y); })
      .attr("font-size", "2em")
      .text(function(d) { return d.y; });
    
      // add the x Axis
    svg.append("g")
        .attr("transform", "translate(0," + height + ")")
        .call(d3.axisBottom(x));

    svg.selectAll(".tick text")
      .attr("font-size", "1.2em")
      .attr("font-weight", "700");

  }
  
  window.addEventListener("beforeunload", function(e){
    var results = [];
    for (let i = 0; i < slidesAITracker.length; i++) {
      for (let j = 0; j < slidesAITracker[i].slides.length; j++) {
        var slideInfo = results.filter(e => e.slide == slidesAITracker[i].slides[j].slide);
        if(slideInfo.length == 0){
          results.push({slide:slidesAITracker[i].slides[j].slide,likes:slidesAITracker[i].slides[j].likes,dislikes:slidesAITracker[i].slides[j].dislikes,messages:slidesAITracker[i].slides[j].messages,question:(slidesAITracker[i].slides[j].question != '' && slidesAITracker[i].slides[j].question == 'YES')? {'YES': 1, 'NO': 0} : (slidesAITracker[i].slides[j].question != '' && slidesAITracker[i].slides[j].question == 'NO') ? {'YES': 0, 'NO': 1} : {'YES': 0, 'NO': 0}});
        }else{
          slideInfo[0].likes += slidesAITracker[i].slides[j].likes;
          slideInfo[0].dislikes += slidesAITracker[i].slides[j].dislikes;
          slideInfo[0].messages.concat(slidesAITracker[i].slides[j].messages);
          if(slidesAITracker[i].slides[j].question != ''){
            if(slidesAITracker[i].slides[j].question == 'YES'){
              slideInfo[0].question['YES']++;
            }else{
              slideInfo[0].question['NO']++;
            }
          }
        }
      }
    }
    results.sort(function(a, b) {
        return b.slide < a.slide ?  1 // if b should come earlier, push a to end
        : b.slide > a.slide ? -1 // if b should come later, push a to begin
        : 0; 
    });
    var jsonData = JSON.stringify(results);
    download(jsonData, 'PresentationResults.json', 'application/json');
  }, false);
}());

`;
}

function clientJSLogic(){
return `
(function() {
  // Don't emit events from inside of notes windows

  window.presentation_type = "client";

  window.presentation_name = "${ powerPointFile.split('.').slice(0, -1).join('.') }";

  window.presentation_expiration = "${ DS_CONFIG.audience_interaction_exp_time }";

  if ( window.location.search.match( /receiver/gi ) ) { return; }

  window.multiplex = Reveal.getConfig().multiplex;

  window.my_socket = io.connect( multiplex.url );

  checkCookie();
  
  var slideAITracker = [];

  function postAIData(data) {

  var messageData = {
    id: getCookie(getCookie(window.presentation_name + "_id")),
    data: data,
    secret: multiplex.secret_c,
    socketId: multiplex.id_c
  };

  my_socket.emit( 'multiplex-statechanged', messageData );

  };

  document.addEventListener("DOMContentLoaded", function(event) {
  document.getElementById("likeButton").addEventListener("click", function() { sendAppreciation(0); });
  document.getElementById("dislikeButton").addEventListener("click", function() { sendAppreciation(1); });
  document.getElementById("messageFormSubmit").addEventListener("submit", function(event) { 
      event.preventDefault();
      sendAppreciation(2,document.getElementById("messageValue").value);
      closeChat(); 
  });
  document.getElementById("messageButton").addEventListener("click", function() { 
      if(!document.getElementById("messageForm").classList.contains("client-chat-open")) {
      openChat();
      }else{
      closeChat();
      }
  });
  document.getElementById("feedbackButton").addEventListener("click", function() {
      if(!document.getElementById("feedbackBoard").classList.contains("client-panel-open")) {
      openFeedback();
      }else{
      closeFeedback();
      }
  });
  });
  
  function sendAppreciation(type,text=null){
      var slideNumber = Reveal.getSlidePastCount();
      var data = {
          type: type,
          slide:slideNumber,
          message: text
      }
      var slideInfo = slideAITracker.filter(e => e.slide == slideNumber);
      if (slideInfo.length == 0) {
      if(type == 2){
          document.getElementById("messageButton").classList.add("disabled");
          closeChat();
      }else{
          slideAITracker.push({slide:slideNumber,type:type})
          document.getElementById("likeButton").classList.add("disabled");
          document.getElementById("dislikeButton").classList.add("disabled");
      }
      }
      postAIData(data);
  }

  function changeAIButtons(){
      var slideNumber = Reveal.getSlidePastCount();
      var slideInfo = slideAITracker.filter(e => e.slide == slideNumber);
      if(slideInfo.length == 0){
      document.getElementById("likeButton").classList.remove("disabled");
      document.getElementById("dislikeButton").classList.remove("disabled");
      document.getElementById("messageButton").classList.remove("disabled")
      }else{
      document.getElementById("likeButton").classList.add("disabled");
      document.getElementById("dislikeButton").classList.add("disabled");
      }
  }

  function openChat() {
  document.getElementById("messageForm").classList.add("client-chat-open");
  }

  function closeChat() {
  document.getElementById("messageForm").classList.remove("client-chat-open");
  }  

  function openFeedback() {
  document.getElementById("feedbackBoard").classList.add("client-panel-open");
  }  

  function closeFeedback() {
  document.getElementById("feedbackBoard").classList.remove("client-panel-open");
  }  

  function getJSON(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', url, true);
  xhr.responseType = 'json';
  xhr.onload = function() {
      var status = xhr.status;
      if (status === 200) {
      callback(null, xhr.response);
      } else {
      callback(status, xhr.response);
      }
  };
  xhr.send();
  };

  function getCookie(cname) {
    var name = cname + "=";
    var decodedCookie = decodeURIComponent(document.cookie);
    var ca = decodedCookie.split(';');
    for(var i = 0; i <ca.length; i++) {
      var c = ca[i];
      while (c.charAt(0) == ' ') {
        c = c.substring(1);
      }
      if (c.indexOf(name) == 0) {
        return c.substring(name.length, c.length);
      }
    }
    return "";
  }

  function setCookie(cname, cvalue, exdays) {
    var d = new Date();
    d.setTime(d.getTime() + (exdays*24*60*60*1000));
    var expires = "expires="+ d.toUTCString();
    document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
  }

  function checkCookie() {
    var prensentation_id = getCookie(window.presentation_name + "_id");
    if (prensentation_id == "") {
      if (prensentation_id != "" && prensentation_id != null) {
        setCookie(window.presentation_name + "_id", makeid(32), parseInt(window.presentation_expiration));
      }
    }
  }

  function makeid(length) {
    var result           = '';
    var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
       result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
 }

  Reveal.addEventListener( 'slidechanged', changeAIButtons);


}());
`
}

function masterAIPanel(){
return ``;

// <div class="client-panel">
//     <svg class="panel-icon" x="0px" y="0px" viewBox="0 0 458 458" style="enable-background:new 0 0 458 458;">
//       <g>
//           <g>
//               <g>
//                   <path d="M428,41.533H30c-16.568,0-30,13.432-30,30v252c0,16.568,13.432,30,30,30h132.1l43.942,52.243
//                       c5.7,6.777,14.103,10.69,22.959,10.69c8.856,0,17.259-3.912,22.959-10.689l43.942-52.243H428c16.569,0,30-13.432,30-30v-252
//                       C458,54.965,444.569,41.533,428,41.533z M428,323.533H281.933L229,386.465l-52.932-62.932H30v-252h398V323.533z"/>
//                   <path d="M85.402,156.999h137c8.284,0,15-6.716,15-15s-6.716-15-15-15h-137c-8.284,0-15,6.716-15,15
//                       S77.118,156.999,85.402,156.999z"/>
//                   <path d="M71,233.999c0,8.284,6.716,15,15,15h286c8.284,0,15-6.716,15-15s-6.716-15-15-15H86
//                       C77.716,218.999,71,225.715,71,233.999z"/>
//               </g>
//           </g>
//       </g>
//     </svg>
//   </div>
}

function clientAIPanel(){
return `
<div id="feedbackBoard" class="client-panel">
  <div id="feedbackButton" class="panel-button"> Feedback </div>
  <svg id="likeButton" class="panel-icon green-svg" x="0px" y="0px" viewBox="0 0 58 58" style="enable-background:new 0 0 58 58;">
  <g>
      <path d="M9.5,43c-2.757,0-5,2.243-5,5s2.243,5,5,5s5-2.243,5-5S12.257,43,9.5,43z"/>
      <path d="M56.5,35c0-2.495-1.375-3.662-2.715-4.233C54.835,29.85,55.5,28.501,55.5,27c0-2.757-2.243-5-5-5H36.134l0.729-3.41
          c0.973-4.549,0.334-9.716,0.116-11.191C36.461,3.906,33.372,0,30.013,0h-0.239C28.178,0,25.5,0.909,25.5,7c0,14.821-6.687,15-7,15
          h0h-1v-2h-16v38h16v-2h28c2.757,0,5-2.243,5-5c0-1.164-0.4-2.236-1.069-3.087C51.745,47.476,53.5,45.439,53.5,43
          c0-1.164-0.4-2.236-1.069-3.087C54.745,39.476,56.5,37.439,56.5,35z M3.5,56V22h12v34H3.5z"/>
  </g>
  </svg>
  <div class="chat-popup" id="messageForm">
  <form id="messageFormSubmit" class="form-container">
      <textarea id="messageValue" placeholder="Type message.." name="msg" required></textarea>
      <button type="submit" class="btn">Send</button>
  </form>
  </div>
  <svg id="messageButton" class="panel-icon" x="0px" y="0px" viewBox="0 0 458 458" style="enable-background:new 0 0 458 458;">
  <g>
      <g>
          <g>
              <path d="M428,41.533H30c-16.568,0-30,13.432-30,30v252c0,16.568,13.432,30,30,30h132.1l43.942,52.243
                  c5.7,6.777,14.103,10.69,22.959,10.69c8.856,0,17.259-3.912,22.959-10.689l43.942-52.243H428c16.569,0,30-13.432,30-30v-252
                  C458,54.965,444.569,41.533,428,41.533z M428,323.533H281.933L229,386.465l-52.932-62.932H30v-252h398V323.533z"/>
              <path d="M85.402,156.999h137c8.284,0,15-6.716,15-15s-6.716-15-15-15h-137c-8.284,0-15,6.716-15,15
                  S77.118,156.999,85.402,156.999z"/>
              <path d="M71,233.999c0,8.284,6.716,15,15,15h286c8.284,0,15-6.716,15-15s-6.716-15-15-15H86
                  C77.716,218.999,71,225.715,71,233.999z"/>
          </g>
      </g>
  </g>
  </svg>
  <svg id="dislikeButton" class="panel-icon red-svg" x="0px" y="0px" viewBox="0 0 58 58" style="enable-background:new 0 0 58 58;" >
  <g>
      <path d="M40.5,0v2h-28c-2.757,0-5,2.243-5,5c0,1.164,0.4,2.236,1.069,3.087C6.255,10.524,4.5,12.561,4.5,15
          c0,1.164,0.4,2.236,1.069,3.087C3.255,18.524,1.5,20.561,1.5,23c0,2.495,1.375,3.662,2.715,4.233C3.165,28.15,2.5,29.499,2.5,31
          c0,2.757,2.243,5,5,5h14.366l-0.729,3.41c-0.973,4.551-0.334,9.717-0.116,11.191C21.539,54.094,24.628,58,27.987,58h0.239
          c1.596,0,4.274-0.909,4.274-7c0-14.82,6.686-15,7-15h0h1v2h16V0H40.5z M54.5,36h-12V2h12V36z"/>
      <path d="M48.5,15c2.757,0,5-2.243,5-5s-2.243-5-5-5s-5,2.243-5,5S45.743,15,48.5,15z"/>
  </g>
  </svg>
</div>
`;
}

function processHTML(str) {
  var div = document.createElement("div");
  div.innerHTML = str.trim();

  return format(div, 0).innerHTML;
}

function format(node, level) {
  var indentBefore = new Array(level++ + 1).join("  "),
    indentAfter = new Array(level - 1).join("  "),
    textNode;

  for (var i = 0; i < node.children.length; i++) {
    textNode = document.createTextNode("\n" + indentBefore);
    node.insertBefore(textNode, node.children[i]);

    format(node.children[i], level);

    if (node.lastElementChild == node.children[i]) {
      textNode = document.createTextNode("\n" + indentAfter);
      node.appendChild(textNode);
    }
  }

  return node;
}

async function processPPTX(pptxFile) {
  var dateBefore = new Date();
  var result = "";
  var myzip = await JSZip.loadAsync(pptxFile.buffer);
  powerPointFile = pptxFile.originalname;
  if (myzip.file("docProps/thumbnail.jpeg") !== null) {
    //var pptxThumbImg = base64ArrayBuffer(await zip.file("docProps/thumbnail.jpeg").async("arraybuffer"));
    //TODO: REVIEW LATER
  }
  var filesInfo = await getContentTypes(myzip);
  slideSize = await getSlideSize(myzip);
  mainthemeContent = await loadTheme(myzip, undefined);

  var numOfSlides = Array.from(Array(filesInfo["slides"].length), (x, i) => i);
  for (const ns of numOfSlides) {
    current_slide = ns;
    var filename = filesInfo["slides"][ns];
    var slide = await processSingleSlide(myzip, filename, ns, slideSize);
    result += slide;
    var progressUpdate = (
      ((parseInt(ns) + 1) * 100) /
      filesInfo["slides"].length
    ).toFixed(0);
    var nextProgress = progressUpdate - currentProgress;
    for (let index = 0; index < nextProgress; index++) {
      process.stdout.write("\u2588");
      cursor++;
    }
    currentProgress += nextProgress;
  }
  process.stdout.write("\n");
  result += genGlobalCSS();
  var dateAfter = new Date();
  console.log("Executio time --> ", dateAfter - dateBefore + "ms");
  return result;
}

async function readXmlFile(zip, filename) {
  var xmlString = await zip.file(filename).async("string");
  return tXml(xmlString);
}

async function getContentTypes(zip) {
  var ContentTypesJson = await readXmlFile(zip, "[Content_Types].xml");
  var subObj = ContentTypesJson["Types"]["Override"];
  var slidesLocArray = [];
  var slideLayoutsLocArray = [];
  for (var i = 0; i < subObj.length; i++) {
    switch (subObj[i]["attrs"]["ContentType"]) {
      case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
        slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
        break;
      case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
        slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
        break;
      default:
    }
  }
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  };
}

async function getSlideSize(zip) {
  // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
  var content = await readXmlFile(zip, "ppt/presentation.xml");
  var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
  return {
    width: (parseInt(sldSzAttrs["cx"]) * 96) / 914400,
    height: (parseInt(sldSzAttrs["cy"]) * 96) / 914400,
  };
}

async function loadTheme(zip, typeFile) {
  if (typeFile !== undefined) {
    var preResContent = await readXmlFile(
      zip,
      typeFile
        .replace("slideMasters/", "slideMasters/_rels/")
        .replace(".xml", ".xml.rels")
    );
  } else {
    var preResContent = await readXmlFile(
      zip,
      "ppt/_rels/presentation.xml.rels"
    );
  }
  var relationshipArray = preResContent["Relationships"]["Relationship"];
  var themeURI = undefined;
  if (relationshipArray.constructor === Array) {
    for (var i = 0; i < relationshipArray.length; i++) {
      if (
        relationshipArray[i]["attrs"]["Type"] ===
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
      ) {
        themeURI = relationshipArray[i]["attrs"]["Target"];
        break;
      }
    }
  } else if (
    relationshipArray["attrs"]["Type"] ===
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  ) {
    themeURI = relationshipArray["attrs"]["Target"];
  }

  if (themeURI === undefined) {
    throw Error("Can't open theme file.");
  }
  return await readXmlFile(zip, "ppt/" + themeURI.replace("../", ""));
}

async function processSingleSlide(zip, sldFileName, index, slideSize) {
  // self.postMessage({
  //     "type": "INFO",
  //     "data": "Processing slide" + (index + 1)
  // });

  // =====< Step 1 >=====
  // Read relationship filename of the slide (Get slideLayoutXX.xml)
  // @sldFileName: ppt/slides/slide1.xml
  // @resName: ppt/slides/_rels/slide1.xml.rels
  var resName =
    sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
  var resContent = await readXmlFile(zip, resName);
  var RelationshipArray = resContent["Relationships"]["Relationship"];
  var layoutFilename = "";
  var slideResObj = {};
  if (RelationshipArray.constructor === Array) {
    for (var i = 0; i < RelationshipArray.length; i++) {
      switch (RelationshipArray[i]["attrs"]["Type"]) {
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
          layoutFilename = RelationshipArray[i]["attrs"]["Target"].replace(
            "../",
            "ppt/"
          );
          break;
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
        default:
          slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
            type: RelationshipArray[i]["attrs"]["Type"].replace(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
              ""
            ),
            target: RelationshipArray[i]["attrs"]["Target"].replace(
              "../",
              "ppt/"
            ),
          };
      }
    }
  } else {
    layoutFilename = RelationshipArray["attrs"]["Target"].replace(
      "../",
      "ppt/"
    );
  }
  // Open slideLayoutXX.xml
  var slideLayoutContent = await readXmlFile(zip, layoutFilename);
  var slideLayoutTables = indexNodes(slideLayoutContent);
  var slideLayoutResObj = {};

  // =====< Step 2 >=====
  // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
  // @resName: ppt/slideLayouts/slideLayout1.xml
  // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
  var slideLayoutResFilename =
    layoutFilename.replace(
      "slideLayouts/slideLayout",
      "slideLayouts/_rels/slideLayout"
    ) + ".rels";
  var slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename);
  RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
  var masterFilename = "";
  if (RelationshipArray.constructor === Array) {
    for (var i = 0; i < RelationshipArray.length; i++) {
      switch (RelationshipArray[i]["attrs"]["Type"]) {
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
          masterFilename = RelationshipArray[i]["attrs"]["Target"].replace(
            "../",
            "ppt/"
          );
          break;
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
        default:
          slideLayoutResObj[RelationshipArray[i]["attrs"]["Id"]] = {
            type: RelationshipArray[i]["attrs"]["Type"].replace(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
              ""
            ),
            target: RelationshipArray[i]["attrs"]["Target"].replace(
              "../",
              "ppt/"
            ),
          };
      }
    }
  } else {
    masterFilename = RelationshipArray["attrs"]["Target"].replace(
      "../",
      "ppt/"
    );
  }
  // Open slideMasterXX.xml
  var slideMasterContent = await readXmlFile(zip, masterFilename);
  var slideMasterTextStyles = getTextByPathList(slideMasterContent, [
    "p:sldMaster",
    "p:txStyles",
  ]);
  var slideMasterTables = indexNodes(slideMasterContent);

  var slideMasterResFilename =
    masterFilename.replace(
      "slideMasters/slideMaster",
      "slideMasters/_rels/slideMaster"
    ) + ".rels";
  var slideMasterResContent = await readXmlFile(zip, slideMasterResFilename);
  RelationshipArray = slideMasterResContent["Relationships"]["Relationship"];
  var masterThemeExists = false;
  var slideMasterResObj = {};
  if (RelationshipArray.constructor === Array) {
    for (var i = 0; i < RelationshipArray.length; i++) {
      switch (RelationshipArray[i]["attrs"]["Type"]) {
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
          masterThemeExists = true;
          break;
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
        default:
          slideMasterResObj[RelationshipArray[i]["attrs"]["Id"]] = {
            type: RelationshipArray[i]["attrs"]["Type"].replace(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
              ""
            ),
            target: RelationshipArray[i]["attrs"]["Target"].replace(
              "../",
              "ppt/"
            ),
          };
      }
    }
  } else {
    masterThemeExists =
      RelationshipArray["attrs"]["Type"] ===
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        ? true
        : false;
  }

  if (masterThemeExists) {
    currentthemeContent = await loadTheme(zip, masterFilename);
  } else {
    currentthemeContent = mainthemeContent;
  }
  var slideContent = await readXmlFile(zip, sldFileName);
  var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
  var nodesMasterContent =
    slideMasterContent["p:sldMaster"]["p:cSld"]["p:spTree"];
  var nodesLayoutContent =
    slideLayoutContent["p:sldLayout"]["p:cSld"]["p:spTree"];
  var warpObj = {
    zip: zip,
    slideLayoutTables: slideLayoutTables,
    slideMasterTables: slideMasterTables,
    slideResObj: slideResObj,
    slideLayoutResObj: slideLayoutResObj,
    slideMasterResObj: slideMasterResObj,
    slideMasterTextStyles: slideMasterTextStyles,
  };

  var bgColor = getSlideBackgroundFill(
    slideContent,
    slideLayoutContent,
    slideMasterContent
  );
  var result =
    "<section style='width:" +
    slideSize.width +
    "px; height:" +
    slideSize.height +
    "px; background-color: #" +
    bgColor +
    ";padding-top: 0px;padding-bottom: 0px;'>";

  if (index > 0) {
    for (var nodeKey in nodesMasterContent) {
      if (nodesMasterContent[nodeKey].constructor === Array) {
        for (var i = 0; i < nodesMasterContent[nodeKey].length; i++) {
          result += await processNodesInSlide(
            nodeKey,
            nodesMasterContent[nodeKey][i],
            warpObj,
            "master"
          );
        }
      } else {
        result += await processNodesInSlide(
          nodeKey,
          nodesMasterContent[nodeKey],
          warpObj,
          "master"
        );
      }
    }
  }

  for (var nodeKey in nodesLayoutContent) {
    if (nodesLayoutContent[nodeKey].constructor === Array) {
      for (var i = 0; i < nodesLayoutContent[nodeKey].length; i++) {
        result += await processNodesInSlide(
          nodeKey,
          nodesLayoutContent[nodeKey][i],
          warpObj,
          "layout"
        );
      }
    } else {
      result += await processNodesInSlide(
        nodeKey,
        nodesLayoutContent[nodeKey],
        warpObj,
        "layout"
      );
    }
  }

  for (var nodeKey in nodes) {
    if (nodes[nodeKey].constructor === Array) {
      for (var i = 0; i < nodes[nodeKey].length; i++) {
        result += await processNodesInSlide(
          nodeKey,
          nodes[nodeKey][i],
          warpObj,
          "simple"
        );
      }
    } else {
      result += await processNodesInSlide(
        nodeKey,
        nodes[nodeKey],
        warpObj,
        "simple"
      );
    }
  }
  return result + "</section>";
}

function indexNodes(content) {
  var keys = Object.keys(content);
  var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

  var idTable = {};
  var idxTable = {};
  var typeTable = {};

  for (var key in spTreeNode) {
    if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
      continue;
    }

    var targetNode = spTreeNode[key];

    if (targetNode.constructor === Array) {
      for (var i = 0; i < targetNode.length; i++) {
        var nvSpPrNode = targetNode[i]["p:nvSpPr"];
        var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
        var idx = getTextByPathList(nvSpPrNode, [
          "p:nvPr",
          "p:ph",
          "attrs",
          "idx",
        ]);
        var type = getTextByPathList(nvSpPrNode, [
          "p:nvPr",
          "p:ph",
          "attrs",
          "type",
        ]);

        if (id !== undefined) {
          idTable[id] = targetNode[i];
        }
        if (idx !== undefined) {
          idxTable[idx] = targetNode[i];
        }
        if (type !== undefined) {
          typeTable[type] = targetNode[i];
        }
      }
    } else {
      var nvSpPrNode = targetNode["p:nvSpPr"];
      var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
      var idx = getTextByPathList(nvSpPrNode, [
        "p:nvPr",
        "p:ph",
        "attrs",
        "idx",
      ]);
      var type = getTextByPathList(nvSpPrNode, [
        "p:nvPr",
        "p:ph",
        "attrs",
        "type",
      ]);

      if (id !== undefined) {
        idTable[id] = targetNode;
      }
      if (idx !== undefined) {
        idxTable[idx] = targetNode;
      }
      if (type !== undefined) {
        typeTable[type] = targetNode;
      }
    }
  }

  return { idTable: idTable, idxTable: idxTable, typeTable: typeTable };
}

async function processNodesInSlide(nodeKey, nodeValue, warpObj, nodeType) {
  var result = "";
  switch (nodeKey) {
    case "p:sp": // Shape, Text
      result = processSpNode(nodeValue, warpObj, nodeType);
      break;
    case "p:cxnSp": // Shape, Text (with connection)
      result = processCxnSpNode(nodeValue, warpObj, nodeType);
      break;
    case "p:pic": // Picture
      result = await processPicNode(nodeValue, warpObj, nodeType);
      break;
    // case "p:graphicFrame": // Chart, Diagram, Table
    //   result = await processGraphicFrameNode(nodeValue, warpObj, nodeType);
    //   break;
    case "p:grpSp": // 群組
      result = processGroupSpNode(nodeValue, warpObj, nodeType);
      break;
    default:
  }

  return result;
}

async function processGroupSpNode(node, warpObj, nodeType) {
  var factor = 96 / 914400;

  var xfrmNode = node["p:grpSpPr"]["a:xfrm"];
  var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * factor;
  var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * factor;
  var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * factor;
  var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * factor;
  var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * factor;
  var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * factor;
  var chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * factor;
  var chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * factor;

  var order = getOrder(node, nodeType);

  var result =
    "<div class='block group' style='z-index: " +
    order +
    "; top: " +
    (y - chy) +
    "px; left: " +
    (x - chx) +
    "px; width: " +
    (cx - chcx) +
    "px; height: " +
    (cy - chcy) +
    "px;'>";

  // Procsee all child nodes
  for (var nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (var i = 0; i < node[nodeKey].length; i++) {
        result += await processNodesInSlide(
          nodeKey,
          node[nodeKey][i],
          warpObj,
          nodeType
        );
      }
    } else {
      result += await processNodesInSlide(
        nodeKey,
        node[nodeKey],
        warpObj,
        nodeType
      );
    }
  }

  result += "</div>";

  return result;
}

function processSpNode(node, warpObj, nodeType) {
  /*
  *  958    <xsd:complexType name="CT_GvmlShape">
  *  959   <xsd:sequence>
  *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
  *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
  *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
  *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
  *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
  *  965   </xsd:sequence>
  *  966 </xsd:complexType>
  */

  var id = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
  var name = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["name"];
  var idx =
    node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined
      ? undefined
      : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
  var type =
    node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined
      ? undefined
      : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
  var order = getOrder(node, nodeType);

  var slideLayoutSpNode = undefined;
  var slideMasterSpNode = undefined;

  if (type !== undefined) {
    if (idx !== undefined) {
      slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
      slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
    } else {
      slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
      slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
    }
  } else {
    if (idx !== undefined) {
      slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
      slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
    } else {
      // Nothing
    }
  }
  if (type === undefined) {
    type = getTextByPathList(slideLayoutSpNode, [
      "p:nvSpPr",
      "p:nvPr",
      "p:ph",
      "attrs",
      "type",
    ]);
    if (type === undefined) {
      type = getTextByPathList(slideMasterSpNode, [
        "p:nvSpPr",
        "p:nvPr",
        "p:ph",
        "attrs",
        "type",
      ]);
    }
  }

  return genShape(
    node,
    slideLayoutSpNode,
    slideMasterSpNode,
    id,
    name,
    idx,
    type,
    order,
    warpObj,
    nodeType
  );
}

function processCxnSpNode(node, warpObj, nodeType) {
  var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
  var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
  //var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
  //var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
  //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
  var order = getOrder(node, nodeType);

  return genShape(
    node,
    undefined,
    undefined,
    id,
    name,
    undefined,
    undefined,
    order,
    warpObj,
    nodeType
  );
}

function genShape(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  id,
  name,
  idx,
  type,
  order,
  warpObj,
  nodeType
) {
  var xfrmList = ["p:spPr", "a:xfrm"];
  var slideXfrmNode = getTextByPathList(node, xfrmList);
  var slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
  var slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);

  var result = "";
  var customG = getTextByPathList(node, ["p:spPr", "a:custGeom"]);
  var shapType = getTextByPathList(node, [
    "p:spPr",
    "a:prstGeom",
    "attrs",
    "prst",
  ]);
  if (
    (nodeType == "master" || nodeType == "layout") &&
    node["p:nvSpPr"] != undefined &&
    node["p:nvSpPr"]["p:cNvSpPr"]["a:spLocks"] != undefined &&
    node["p:nvSpPr"]["p:cNvSpPr"]["a:spLocks"]["attrs"]["noGrp"] == "1"
  ) {
    return "";
  }
  if (
    (nodeType == "master" || nodeType == "layout") &&
    customG === undefined &&
    shapType === undefined
  ) {
    return "";
  }
  var isFlipV = getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1";
  var isFlipH = getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1";
  if (shapType !== undefined) {
    var off = getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
    var x = (parseInt(off["x"]) * 96) / 914400;
    var y = (parseInt(off["y"]) * 96) / 914400;

    var ext = getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
    var w = (parseInt(ext["cx"]) * 96) / 914400;
    var h =
      (parseInt(ext["cy"]) * 96) / 914400 == 0
        ? 1
        : (parseInt(ext["cy"]) * 96) / 914400;

    var rot = getTextByPathList(slideXfrmNode, ["attrs", "rot"]);
    if (isFlipH && isFlipV) {
      var flip = `transform="scale(1, -1)`; //This shouldnt be like this??? REVIEW
    } else if (isFlipH) {
      var flip = `transform="scale(-1, 1)`;
    } else if (isFlipV) {
      var flip = `transform="scale(1, -1)`;
    } else {
      var flip = "";
    }

    if (rot != undefined) {
      if (flip != "") {
        flip += ` rotate(${parseInt(rot) * (3 / 180000)})"`;
      } else {
        flip = `transform="rotate(${parseInt(rot) * (3 / 180000)})"`;
      }
    } else {
      if (flip != "") flip += `"`;
    }

    result +=
      "<svg class='drawing' _id='" +
      id +
      "' _idx='" +
      idx +
      "' _type='" +
      type +
      "' _name='" +
      name +
      "' style='" +
      getPosition(slideXfrmNode, undefined, undefined) +
      getSize(slideXfrmNode, undefined, undefined) +
      " z-index: " +
      order +
      ";" +
      "' " +
      flip +
      ">";

    // Fill Color
    var fillColor = getShapeFill(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      true
    );

    // Border Color
    var border = getBorder(node, slideLayoutSpNode, slideMasterSpNode, true);

    var headEndNodeAttrs = getTextByPathList(node, [
      "p:spPr",
      "a:ln",
      "a:headEnd",
      "attrs",
    ]);
    var tailEndNodeAttrs = getTextByPathList(node, [
      "p:spPr",
      "a:ln",
      "a:tailEnd",
      "attrs",
    ]);
    // type: none, triangle, stealth, diamond, oval, arrow
    if (
      (headEndNodeAttrs !== undefined &&
        (headEndNodeAttrs["type"] === "triangle" ||
          headEndNodeAttrs["type"] === "arrow")) ||
      (tailEndNodeAttrs !== undefined &&
        (tailEndNodeAttrs["type"] === "triangle" ||
          tailEndNodeAttrs["type"] === "arrow"))
    ) {
      var triangleMarker = `<defs><marker id="markerTriangle${currentProgress}" viewBox="0 0 10 10" refX="1" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse" markerUnits="strokeWidth"><path fill="${border.color}" d="M 0 0 L 10 5 L 0 10 z" /></marker></defs>`;
      result += triangleMarker;
    }

    switch (shapType) {
      case "accentBorderCallout1":
      case "accentBorderCallout2":
      case "accentBorderCallout3":
      case "accentCallout1":
      case "accentCallout2":
      case "accentCallout3":
      case "actionButtonBackPrevious":
      case "actionButtonBeginning":
      case "actionButtonBlank":
      case "actionButtonDocument":
      case "actionButtonEnd":
      case "actionButtonForwardNext":
      case "actionButtonHelp":
      case "actionButtonHome":
      case "actionButtonInformation":
      case "actionButtonMovie":
      case "actionButtonReturn":
      case "actionButtonSound":
      case "arc":
      case "bevel":
      case "blockArc":
      case "borderCallout1":
      case "borderCallout2":
      case "borderCallout3":
      case "bracePair":
      case "bracketPair":
      case "callout1":
      case "callout2":
      case "callout3":
      case "can":
      case "chartPlus":
      case "chartStar":
      case "chartX":
      case "chevron":
      case "chord":
      case "cloud":
      case "cloudCallout":
      case "corner":
      case "cornerTabs":
      case "cube":
      case "decagon":
      case "diagStripe":
      case "diamond":
      case "dodecagon":
      case "donut":
      case "doubleWave":
      case "downArrowCallout":
      case "ellipseRibbon":
      case "ellipseRibbon2":
      case "flowChartAlternateProcess":
      case "flowChartCollate":
      case "flowChartConnector":
      case "flowChartDecision":
      case "flowChartDelay":
      case "flowChartDisplay":
      case "flowChartDocument":
      case "flowChartExtract":
      case "flowChartInputOutput":
      case "flowChartInternalStorage":
      case "flowChartMagneticDisk":
      case "flowChartMagneticDrum":
      case "flowChartMagneticTape":
      case "flowChartManualInput":
      case "flowChartManualOperation":
      case "flowChartMerge":
      case "flowChartMultidocument":
      case "flowChartOfflineStorage":
      case "flowChartOffpageConnector":
      case "flowChartOnlineStorage":
      case "flowChartOr":
      case "flowChartPredefinedProcess":
      case "flowChartPreparation":
      case "flowChartProcess":
      case "flowChartPunchedCard":
      case "flowChartPunchedTape":
      case "flowChartSort":
      case "flowChartSummingJunction":
      case "flowChartTerminator":
      case "folderCorner":
      case "frame":
      case "funnel":
      case "gear6":
      case "gear9":
      case "halfFrame":
      case "heart":
      case "heptagon":
      case "hexagon":
      case "homePlate":
      case "horizontalScroll":
      case "irregularSeal1":
      case "irregularSeal2":
      case "leftArrow":
      case "leftArrowCallout":
      case "leftBrace":
      case "leftBracket":
      case "leftRightArrowCallout":
      case "leftRightRibbon":
      case "irregularSeal1":
      case "lightningBolt":
      case "lineInv":
      case "mathDivide":
      case "mathEqual":
      case "mathMinus":
      case "mathMultiply":
      case "mathNotEqual":
      case "mathPlus":
      case "moon":
      case "nonIsoscelesTrapezoid":
      case "noSmoking":
      case "octagon":
      case "parallelogram":
      case "pentagon":
      case "pie":
      case "pieWedge":
      case "plaque":
      case "plaqueTabs":
      case "plus":
      case "quadArrowCallout":
      case "ribbon":
      case "ribbon2":
      case "rightArrowCallout":
      case "rightBrace":
      case "rightBracket":
      case "round1Rect":
      case "round2DiagRect":
      case "round2SameRect":
      case "rtTriangle":
      case "smileyFace":
      case "snip1Rect":
      case "snip2DiagRect":
      case "snip2SameRect":
      case "snipRoundRect":
      case "squareTabs":
      case "star10":
      case "star12":
      case "star16":
      case "star24":
      case "star32":
      case "star4":
      case "star5":
      case "star6":
      case "star7":
      case "star8":
      case "sun":
      case "teardrop":
      case "trapezoid":
      case "upArrowCallout":
      case "upDownArrowCallout":
      case "verticalScroll":
      case "wave":
      case "wedgeEllipseCallout":
      case "wedgeRectCallout":
      case "wedgeRoundRectCallout":
      case "rect":
        result +=
          "<rect x='0' y='0' width='" +
          w +
          "' height='" +
          h +
          "' fill='" +
          fillColor +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          border.width +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' />";
        break;
      case "ellipse":
        result +=
          "<ellipse cx='" +
          w / 2 +
          "' cy='" +
          h / 2 +
          "' rx='" +
          w / 2 +
          "' ry='" +
          h / 2 +
          "' fill='" +
          fillColor +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          border.width +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' />";
        break;
      case "roundRect":
        result +=
          "<rect x='0' y='0' width='" +
          w +
          "' height='" +
          h +
          "' rx='7' ry='7' fill='" +
          fillColor +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          border.width +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' />";
        break;
      case "bentConnector2": // 直角 (path)
        var d = "";
        if (isFlipV) {
          d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
        } else {
          d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
        }
        result +=
          "<path d='" +
          d +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          border.width +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' fill='none' ";
        if (
          headEndNodeAttrs !== undefined &&
          (headEndNodeAttrs["type"] === "triangle" ||
            headEndNodeAttrs["type"] === "arrow")
        ) {
          result +=
            "marker-start='url(#markerTriangle" + currentProgress + ")' ";
        }
        if (
          tailEndNodeAttrs !== undefined &&
          (tailEndNodeAttrs["type"] === "triangle" ||
            tailEndNodeAttrs["type"] === "arrow")
        ) {
          result += "marker-end='url(#markerTriangle" + currentProgress + ")' ";
        }
        result += "/>";
        break;
      case "line":
      case "straightConnector1":
      case "bentConnector3":
      case "bentConnector4":
      case "bentConnector5":
      case "curvedConnector2":
      case "curvedConnector3":
      case "curvedConnector4":
      case "curvedConnector5":
        if (isFlipV) {
          result +=
            "<line x1='" +
            w +
            "' y1='0' x2='0' y2='" +
            h +
            "' stroke='" +
            border.color +
            "' stroke-width='" +
            border.width +
            "' stroke-dasharray='" +
            border.strokeDasharray +
            "' ";
        } else {
          result +=
            "<line x1='0' y1='0' x2='" +
            w +
            "' y2='" +
            h +
            "' stroke='" +
            border.color +
            "' stroke-width='" +
            border.width +
            "' stroke-dasharray='" +
            border.strokeDasharray +
            "' ";
        }
        if (
          headEndNodeAttrs !== undefined &&
          (headEndNodeAttrs["type"] === "triangle" ||
            headEndNodeAttrs["type"] === "arrow")
        ) {
          result +=
            "marker-start='url(#markerTriangle" + currentProgress + ")' ";
        }
        if (
          tailEndNodeAttrs !== undefined &&
          (tailEndNodeAttrs["type"] === "triangle" ||
            tailEndNodeAttrs["type"] === "arrow")
        ) {
          result += "marker-end='url(#markerTriangle" + currentProgress + ")' ";
        }
        result += "/>";
        break;
      case "rightArrow":
        result +=
          '<defs><marker id="markerTriangle" viewBox="0 0 10 10" refX="1" refY="5" markerWidth="2.5" markerHeight="2.5" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M 0 0 L 10 5 L 0 10 z" /></marker></defs>';
        result +=
          "<line x1='0' y1='" +
          h / 2 +
          "' x2='" +
          (w - 15) +
          "' y2='" +
          h / 2 +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          h / 2 +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' ";
        result += "marker-end='url(#markerTriangle" + currentProgress + ")' />";
        break;
      case "downArrow":
        result +=
          '<defs><marker id="markerTriangle" viewBox="0 0 10 10" refX="1" refY="5" markerWidth="2.5" markerHeight="2.5" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M 0 0 L 10 5 L 0 10 z" /></marker></defs>';
        result +=
          "<line x1='" +
          w / 2 +
          "' y1='0' x2='" +
          w / 2 +
          "' y2='" +
          (h - 15) +
          "' stroke='" +
          border.color +
          "' stroke-width='" +
          w / 2 +
          "' stroke-dasharray='" +
          border.strokeDasharray +
          "' ";
        result += "marker-end='url(#markerTriangle" + currentProgress + ")' />";
        break;
      case "bentArrow":
      case "bentUpArrow":
      case "stripedRightArrow":
      case "quadArrow":
      case "circularArrow":
      case "swooshArrow":
      case "leftRightArrow":
      case "leftRightUpArrow":
      case "leftUpArrow":
      case "leftCircularArrow":
      case "notchedRightArrow":
      case "curvedDownArrow":
      case "curvedLeftArrow":
      case "curvedRightArrow":
      case "curvedUpArrow":
      case "upDownArrow":
      case "upArrow":
      case "uturnArrow":
      case "leftRightCircularArrow":
        break;
      case "triangle":
        break;
      case undefined:
      default:
        log_info += "Undefine shape type." + "\n";
    }

    result += "</svg>";

    result +=
      "<div class='block content " +
      getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
      "' _id='" +
      id +
      "' _idx='" +
      idx +
      "' _type='" +
      type +
      "' _name='" +
      name +
      "' style='" +
      getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getTextOrientation(node, slideLayoutSpNode, slideMasterSpNode) +
      getPadding(node, slideLayoutSpNode, slideMasterSpNode, type) +
      // getBorder(node, slideLayoutSpNode, slideMasterSpNode, false) +
      " z-index: " +
      order +
      ";" +
      "'>";

    // TextBody
    if (node["p:txBody"] !== undefined) {
      result += genTextBody(
        node,
        slideLayoutSpNode,
        slideMasterSpNode,
        type,
        warpObj
      );
    }

    result += "</div>";
  } else if (nodeType == "master" || nodeType == "layout") {
    var off = getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
    var x = (parseInt(off["x"]) * 96) / 914400;
    var y = (parseInt(off["y"]) * 96) / 914400;

    var ext = getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
    var w = (parseInt(ext["cx"]) * 96) / 914400;
    var h = (parseInt(ext["cy"]) * 96) / 914400;

    result +=
      "<svg class='drawing' _id='" +
      id +
      "' _idx='" +
      idx +
      "' _type='" +
      type +
      "' _name='" +
      name +
      "' style='" +
      getPosition(slideXfrmNode, undefined, undefined) +
      getSize(slideXfrmNode, undefined, undefined) +
      " z-index: " +
      order +
      ";" +
      "'>";

    // Fill Color
    var fillColor = getShapeFill(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      true
    );

    // Border Color
    var border = getBorder(node, slideLayoutSpNode, slideMasterSpNode, true);

    var headEndNodeAttrs = getTextByPathList(node, [
      "p:spPr",
      "a:ln",
      "a:headEnd",
      "attrs",
    ]);
    var tailEndNodeAttrs = getTextByPathList(node, [
      "p:spPr",
      "a:ln",
      "a:tailEnd",
      "attrs",
    ]);
    // type: none, triangle, stealth, diamond, oval, arrow
    if (
      (headEndNodeAttrs !== undefined &&
        (headEndNodeAttrs["type"] === "triangle" ||
          headEndNodeAttrs["type"] === "arrow")) ||
      (tailEndNodeAttrs !== undefined &&
        (tailEndNodeAttrs["type"] === "triangle" ||
          tailEndNodeAttrs["type"] === "arrow"))
    ) {
      var triangleMarker = `<defs><marker id="markerTriangle${currentProgress}" viewBox="0 0 10 10" refX="1" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse" markerUnits="strokeWidth"><path fill="${border.color}" d="M 0 0 L 10 5 L 0 10 z" /></marker></defs>`;
      result += triangleMarker;
    }
    for (var k in customG["a:pathLst"]) {
      if (k == "attrs") {
        continue;
      }
      var le_path = Object.values(customG["a:pathLst"][k]);
      var le_path_keys = Object.keys(customG["a:pathLst"][k]);
      var listPath = [];
      for (let k = 0; k < le_path_keys.length - 1; k++) {
        if (!Array.isArray(le_path[k])) {
          le_path[k] = [le_path[k]];
        }
        for (let j = 0; j < le_path[k].length; j++) {
          le_path[k][j]["type"] = le_path_keys[k];
          listPath.push(le_path[k][j]);
        }
      }
      var le_path_dim = le_path[le_path.length - 1];
      listPath.sort((a, b) =>
        a["attrs"]["order"] > b["attrs"]["order"] ? 1 : -1
      );
      result += `<path d="`;
      for (var i = 0; i < listPath.length; i++) {
        switch (listPath[i]["type"]) {
          case "a:moveTo":
            result += `M ${
              parseInt(listPath[i]["a:pt"]["attrs"]["x"]) *
              (w / parseInt(le_path_dim["w"]))
            } ${
              parseInt(listPath[i]["a:pt"]["attrs"]["y"]) *
              (h / parseInt(le_path_dim["h"]))
            } `;
            break;
          case "a:lnTo":
            result += `L ${
              parseInt(listPath[i]["a:pt"]["attrs"]["x"]) *
              (w / parseInt(le_path_dim["w"]))
            } ${
              parseInt(listPath[i]["a:pt"]["attrs"]["y"]) *
              (h / parseInt(le_path_dim["h"]))
            } `;
            break;
          case "a:cubicBezTo":
            result += `C ${
              parseInt(listPath[i]["a:pt"][0]["attrs"]["x"]) *
              (w / parseInt(le_path_dim["w"]))
            } ${
              parseInt(listPath[i]["a:pt"][0]["attrs"]["y"]) *
              (h / parseInt(le_path_dim["h"]))
            }, ${
              parseInt(listPath[i]["a:pt"][1]["attrs"]["x"]) *
              (w / parseInt(le_path_dim["w"]))
            } ${
              parseInt(listPath[i]["a:pt"][1]["attrs"]["y"]) *
              (h / parseInt(le_path_dim["h"]))
            }, ${
              parseInt(listPath[i]["a:pt"][2]["attrs"]["x"]) *
              (w / parseInt(le_path_dim["w"]))
            } ${
              parseInt(listPath[i]["a:pt"][2]["attrs"]["y"]) *
              (h / parseInt(le_path_dim["h"]))
            } `;
            break;
          case "a:close":
            break;
          default:
            log_info += "Svg error - " + listPath[i] + "\n";
        }
      }
      if (isFlipH && isFlipV) {
        var flip = `transform="scale(-1, -1) translate(-${w}, -${h})"`;
      } else if (isFlipH) {
        var flip = `transform="scale(1, -1) translate(0, -${h})"`;
      } else if (isFlipV) {
        var flip = `transform="scale(-1, 1) translate(-${w},0)"`;
      } else {
        var flip = "";
      }
      result += `" ${flip} stroke="${border.color}" stroke-width="${border.width}" stroke-dasharray="${border.strokeDasharray}" fill="${fillColor}"/>`;
    }

    result += "</svg>";

    result +=
      "<div class='block content " +
      getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
      "' _id='" +
      id +
      "' _idx='" +
      idx +
      "' _type='" +
      type +
      "' _name='" +
      name +
      "' style='" +
      getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getBorder(node, slideLayoutSpNode, slideMasterSpNode, false) +
      getPadding(node, slideLayoutSpNode, slideMasterSpNode, type) +
      getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getTextOrientation(node, slideLayoutSpNode, slideMasterSpNode) +
      " z-index: " +
      order +
      ";" +
      "'>";

    // TextBody
    if (node["p:txBody"] !== undefined) {
      result += genTextBody(
        node,
        slideLayoutSpNode,
        slideMasterSpNode,
        type,
        warpObj
      );
    }
    result += "</div>";
  } else {
    result +=
      "<div class='block content " +
      getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
      "' _id='" +
      id +
      "' _idx='" +
      idx +
      "' _type='" +
      type +
      "' _name='" +
      name +
      "' style='" +
      getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getPadding(node, slideLayoutSpNode, slideMasterSpNode, type) +
      getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      getBorder(node, slideLayoutSpNode, slideMasterSpNode, false) +
      getShapeFill(node, slideLayoutSpNode, slideMasterSpNode, false) +
      getTextOrientation(node, slideLayoutSpNode, slideMasterSpNode) +
      " z-index: " +
      order +
      ";" +
      "'>";

    // TextBody
    if (node["p:txBody"] !== undefined) {
      result += genTextBody(
        node,
        slideLayoutSpNode,
        slideMasterSpNode,
        type,
        warpObj
      );
    }
    result += "</div>";
  }

  return result;
}

async function processPicNode(node, warpObj, nodeType) {
  var order = getOrder(node, nodeType);
  var rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
  if (nodeType == "simple") {
    var imgName = warpObj["slideResObj"][rid]["target"];
  } else if (nodeType == "layout") {
    var imgName = warpObj["slideLayoutResObj"][rid]["target"];
  } else {
    var imgName = warpObj["slideMasterResObj"][rid]["target"];
  }
  var imgFileExt = extractFileExtension(imgName).toLowerCase();
  var zip = warpObj["zip"];

  var imageData = await zip.file(imgName).async("nodebuffer");

  var image = sharp(imageData);

  var imageMD = await image.metadata();

  // var imgArrayBuffer = zip.file(imgName).asArrayBuffer();
  var mimeType = "";
  var xfrmNode = node["p:spPr"]["a:xfrm"];

  var croppingNode = node["p:blipFill"]["a:srcRect"];
  var strechingNode = node["p:blipFill"]["a:stretch"];
  var tilingNode = node["p:blipFill"]["a:tile"];
  var result_width = parseInt(imageMD.width);
  var result_height = parseInt(imageMD.height);

  switch (imgFileExt) {
    case "jpg":
    case "jpeg":
      mimeType = "image/jpeg";
      break;
    case "png":
      mimeType = "image/png";
      break;
    case "gif":
      mimeType = "image/gif";
      break;
    case "emf": // Not native support
      mimeType = "image/x-emf";
      break;
    case "wmf": // Not native support
      mimeType = "image/x-wmf";
      break;
    default:
      mimeType = "image/*";
  }

  if (croppingNode != undefined) {
    var top =
      croppingNode["attrs"]["t"] == undefined
        ? 0
        : parseInt(croppingNode["attrs"]["t"]) / 100000;
    var bottom =
      croppingNode["attrs"]["b"] == undefined
        ? 0
        : parseInt(croppingNode["attrs"]["b"]) / 100000;
    var left =
      croppingNode["attrs"]["l"] == undefined
        ? 0
        : parseInt(croppingNode["attrs"]["l"]) / 100000;
    var right =
      croppingNode["attrs"]["r"] == undefined
        ? 0
        : parseInt(croppingNode["attrs"]["r"]) / 100000;

    image.extract({
      left: Math.round(result_width * left, 0),
      top: Math.round(result_height * top, 0),
      width:
        Math.round(result_width * (1 - right), 0) -
        Math.round(result_width * left, 0),
      height:
        Math.round(result_height * (1 - bottom), 0) -
        Math.round(result_height * top, 0),
    });
  }
  if (strechingNode != undefined) {
    if (strechingNode["a:fillRect"] != undefined) {
      result_width = "100%";
      result_height = "100%";
    }
  }
  if (tilingNode != undefined) {
  }

  try {
    var imageBuffer = await image.toBuffer();
    var image_for_effects = await Jimp.read(imageBuffer);
    applyImageEffects(node, image_for_effects, xfrmNode, mimeType);
  
    var imageB64 = await image_for_effects.getBase64Async(mimeType);
    return (
      "<div class='block content' style='" +
      getPosition(xfrmNode, undefined, undefined) +
      getSize(xfrmNode, undefined, undefined) +
      " z-index: " +
      order +
      ";" +
      "'><img src=\" " +
      imageB64 +
      "\" style='width: " +
      "100%" +
      "; height: " +
      "100%" +
      ";'/></div>"
    );
  } catch (e) {
      console.error("There was an error processing an image on slide " + current_slide+1 + " and the image could not be processed!");
      return (
        "<div class='block content' style='" +
        getPosition(xfrmNode, undefined, undefined) +
        getSize(xfrmNode, undefined, undefined) +
        " z-index: " +
        order +
        ";" +
        "'><img src=\" " +
        "\" style='width: " +
        "100%" +
        "; height: " +
        "100%" +
        ";'/></div>"
      );
  }

}

async function processGraphicFrameNode(node, warpObj, nodeType) {
  var result = "";
  var graphicTypeUri = getTextByPathList(node, [
    "a:graphic",
    "a:graphicData",
    "attrs",
    "uri",
  ]);

  switch (graphicTypeUri) {
    case "http://schemas.openxmlformats.org/drawingml/2006/table":
      result = genTable(node, warpObj, nodeType);
      break;
    case "http://schemas.openxmlformats.org/drawingml/2006/chart":
      result = await genChart(node, warpObj, nodeType);
      break;
    case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
      result = genDiagram(node, warpObj, nodeType);
      break;
    default:
  }

  return result;
}

function processSpPrNode(node, warpObj) {
  /*
  * 2241 <xsd:complexType name="CT_ShapeProperties">
  * 2242   <xsd:sequence>
  * 2243     <xsd:element name="xfrm" type="CT_Transform2D"  minOccurs="0" maxOccurs="1"/>
  * 2244     <xsd:group   ref="EG_Geometry"                  minOccurs="0" maxOccurs="1"/>
  * 2245     <xsd:group   ref="EG_FillProperties"            minOccurs="0" maxOccurs="1"/>
  * 2246     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
  * 2247     <xsd:group   ref="EG_EffectProperties"          minOccurs="0" maxOccurs="1"/>
  * 2248     <xsd:element name="scene3d" type="CT_Scene3D"   minOccurs="0" maxOccurs="1"/>
  * 2249     <xsd:element name="sp3d" type="CT_Shape3D"      minOccurs="0" maxOccurs="1"/>
  * 2250     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
  * 2251   </xsd:sequence>
  * 2252   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
  * 2253 </xsd:complexType>
  */
  // TODO:
}

function genTextBody(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  warpObj
) {
  var text = "";
  var text_complete_string = "";
  var id = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
  var tag = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["descr"];
  var textBodyNode = node["p:txBody"];
  var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
  if (textBodyNode === undefined) {
    return text;
  }

  if (textBodyNode["a:p"].constructor === Array) {
    // multi p
    for (var i = 0; i < textBodyNode["a:p"].length; i++) {
      var pNode = textBodyNode["a:p"][i];
      var rNode = pNode["a:r"];
      text +=
        "<div class='" +
        getHorizontalAlign(
          pNode,
          slideLayoutSpNode,
          slideMasterSpNode,
          type,
          slideMasterTextStyles
        ) +
        "'>";
      text += genBuChar(
        pNode,
        slideLayoutSpNode,
        slideMasterSpNode,
        type,
        slideMasterTextStyles
      );
      if (rNode === undefined) {
        // without r
        var lvl_string =
          getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
            ? "a:lvl1pPr"
            : "a:lvl" +
              (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
              "pPr";
        text += genSpanElement(
          pNode,
          slideLayoutSpNode,
          slideMasterSpNode,
          id,
          type,
          warpObj,
          lvl_string
        );
        text_complete_string += getTextFromNode(pNode);
      } else if (rNode.constructor === Array) {
        // with multi r
        for (var j = 0; j < rNode.length; j++) {
          var lvl_string =
            getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
              ? "a:lvl1pPr"
              : "a:lvl" +
                (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
                "pPr";
          text += genSpanElement(
            rNode[j],
            slideLayoutSpNode,
            slideMasterSpNode,
            id,
            type,
            warpObj,
            lvl_string
          );
          text_complete_string += getTextFromNode(rNode[j]);
        }
      } else {
        // with one r
        var lvl_string =
          getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
            ? "a:lvl1pPr"
            : "a:lvl" +
              (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
              "pPr";
        text += genSpanElement(
          rNode,
          slideLayoutSpNode,
          slideMasterSpNode,
          id,
          type,
          warpObj,
          lvl_string
        );
        text_complete_string += getTextFromNode(rNode);
      }
      text_complete_string += "\n";
      text += "</div>";
    }
  } else {
    // one p
    var pNode = textBodyNode["a:p"];
    var rNode = pNode["a:r"];
    text +=
      "<div class='" +
      getHorizontalAlign(
        pNode,
        slideLayoutSpNode,
        slideMasterSpNode,
        type,
        slideMasterTextStyles
      ) +
      "'>";
    text += genBuChar(
      pNode,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideMasterTextStyles
    );
    if (rNode === undefined) {
      // without r
      var lvl_string =
        getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
          ? "a:lvl1pPr"
          : "a:lvl" +
            (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
            "pPr";
      text += genSpanElement(
        pNode,
        slideLayoutSpNode,
        slideMasterSpNode,
        id,
        type,
        warpObj,
        lvl_string
      );
      text_complete_string += getTextFromNode(pNode);
    } else if (rNode.constructor === Array) {
      // with multi r
      for (var j = 0; j < rNode.length; j++) {
        var lvl_string =
          getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
            ? "a:lvl1pPr"
            : "a:lvl" +
              (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
              "pPr";
        text += genSpanElement(
          rNode[j],
          slideLayoutSpNode,
          slideMasterSpNode,
          id,
          type,
          warpObj,
          lvl_string
        );
        text_complete_string += getTextFromNode(rNode[j]);
      }
    } else {
      // with one r
      var lvl_string =
        getTextByPathStr(pNode, "a:pPr attrs lvl") === undefined
          ? "a:lvl1pPr"
          : "a:lvl" +
            (parseInt(getTextByPathStr(pNode, "a:pPr attrs lvl")) + 1) +
            "pPr";
      text += genSpanElement(
        rNode,
        slideLayoutSpNode,
        slideMasterSpNode,
        id,
        type,
        warpObj,
        lvl_string
      );
      text_complete_string += getTextFromNode(rNode);
    }
    text_complete_string += "\n";
    text += "</div>";
  }
  if (tag != undefined) {
    return evaluateCommandText(text_complete_string, tag);
  } else {
    return text;
  }
}

function genBuChar(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideMasterTextStyles
) {
  var pPrNode = node["a:pPr"];

  var lvl = parseInt(getTextByPathList(pPrNode, ["attrs", "lvl"]));
  if (isNaN(lvl)) {
    lvl = 0;
  }

  var buChar = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
  if (buChar === undefined) {
    buChar = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:p",
      "a:pPr",
      "a:buChar",
      "attrs",
      "char",
    ]);
    pPrNode =
      buChar !== undefined
        ? slideLayoutSpNode["p:txBody"]["a:p"]["a:pPr"]
        : undefined;
    if (buChar === undefined) {
      buChar = getTextByPathList(slideMasterSpNode, [
        "p:txBody",
        "a:p",
        "a:pPr",
        "a:buChar",
        "attrs",
        "char",
      ]);
      pPrNode =
        buChar !== undefined
          ? slideMasterSpNode["p:txBody"]["a:p"]["a:pPr"]
          : undefined;
      if (buChar === undefined) {
        switch (type) {
          case "title":
          case "subTitle":
          case "ctrTitle":
            buChar = getTextByPathList(slideMasterTextStyles, [
              "p:titleStyle",
              "a:lvl1pPr",
              "a:buChar",
              "attrs",
              "char",
            ]);
            pPrNode =
              buChar !== undefined
                ? slideMasterTextStyles["p:bodyStyle"]["a:lvl1pPr"]
                : undefined;
            break;
          case "body":
            buChar = getTextByPathList(slideMasterTextStyles, [
              "p:bodyStyle",
              "a:lvl1pPr",
              "a:buChar",
              "attrs",
              "char",
            ]);
            pPrNode =
              buChar !== undefined
                ? slideMasterTextStyles["p:bodyStyle"]["a:lvl1pPr"]
                : undefined;
            break;
        }
      }
    }
  }

  if (buChar !== undefined) {
    var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
    if (buFontAttrs !== undefined) {
      var marginLeft =
        (parseInt(getTextByPathList(pPrNode, ["attrs", "marL"])) * 96) / 914400;
      var marginRight = parseInt(buFontAttrs["pitchFamily"]);
      if (isNaN(marginLeft)) {
        marginLeft = (328600 * 96) / 914400;
      }
      if (isNaN(marginRight)) {
        marginRight = 0;
      }
      var typeface = buFontAttrs["typeface"];

      return (
        "<span style='font-family: " +
        typeface +
        "; margin-left: " +
        marginLeft * lvl +
        "px" +
        "; margin-right: " +
        marginRight +
        "px" +
        "; font-size: 20pt" +
        "'>" +
        buChar +
        "</span>"
      );
    } else {
      marginLeft = ((328600 * 96) / 914400) * lvl;
      return (
        "<span style='margin-left: " + marginLeft + "px;'>" + buChar + "</span>"
      );
    }
  } else {
    //buChar = '•';
    return (
      "<span style='margin-left: " +
      ((328600 * 96) / 914400) * lvl +
      "px" +
      "; margin-right: " +
      0 +
      "px;'></span>"
    );
  }

  return "";
}

function genSpanElement(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  id,
  type,
  warpObj,
  lvl
) {
  var slideLayoutTables = warpObj["slideLayoutTables"];
  var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

  var text = node["a:t"];
  if (typeof text !== "string") {
    text = getTextByPathList(node, ["a:fld", "a:t"]);
    if (typeof text !== "string") {
      text = "&nbsp;";
    }
  }
  var styleText =
    "color:" +
    getFontColor(
      node,
      id,
      type,
      slideLayoutSpNode,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";font-size:" +
    getFontSize(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";margin:" +
    getMarginsSize(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";text-indent:" +
    getIndentSize(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";text-transform:" +
    getTextCap(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";font-family:" +
    getFontType(node, type, slideMasterTextStyles) +
    ";font-weight:" +
    getFontBold(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";font-style:" +
    getFontItalic(
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type,
      slideLayoutTables,
      slideMasterTextStyles,
      lvl
    ) +
    ";line-height:" +
    getLineSpacing(node, type, slideMasterTextStyles) +
    ";text-decoration:" +
    getFontDecoration(node, type, slideMasterTextStyles) +
    ";vertical-align:" +
    getTextVerticalAlign(node, type, slideMasterTextStyles) +
    ";";

  var cssName = "";

  if (styleText in styleTable) {
    cssName = styleTable[styleText]["name"];
  } else {
    cssName = "_css_" + (Object.keys(styleTable).length + 1);
    styleTable[styleText] = {
      name: cssName,
      text: styleText,
    };
  }

  var linkID = getTextByPathList(node, [
    "a:rPr",
    "a:hlinkClick",
    "attrs",
    "r:id",
  ]);
  if (linkID !== undefined) {
    var linkURL = warpObj["slideResObj"][linkID]["target"];
    return (
      "<span class='text-block " +
      cssName +
      "'><a href='" +
      linkURL +
      "' target='_blank'>" +
      text.replace(/\s/i, "&nbsp;") +
      "</a></span>"
    );
  } else {
    return (
      "<span class='text-block " +
      cssName +
      "'>" +
      text.replace(/\s/i, "&nbsp;") +
      "</span>"
    );
  }
}

function genGlobalCSS() {
  var cssText = "";
  for (var key in styleTable) {
    cssText +=
      "section ." +
      styleTable[key]["name"] +
      "{" +
      styleTable[key]["text"] +
      "}\n";
  }
  return "<style>" + cssText + "</style>";
}

function genTable(node, warpObj, nodeType) {
  var order = getOrder(node, nodeType);
  var tableNode = getTextByPathList(node, [
    "a:graphic",
    "a:graphicData",
    "a:tbl",
  ]);
  var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
  var tableHtml =
    "<table style='" +
    getPosition(xfrmNode, undefined, undefined) +
    getSize(xfrmNode, undefined, undefined) +
    " z-index: " +
    order +
    ";'>";

  var trNodes = tableNode["a:tr"];
  if (trNodes.constructor === Array) {
    for (var i = 0; i < trNodes.length; i++) {
      tableHtml += "<tr>";
      var tcNodes = trNodes[i]["a:tc"];

      if (tcNodes.constructor === Array) {
        for (var j = 0; j < tcNodes.length; j++) {
          var text = genTextBody(
            tcNodes[j],
            undefined,
            undefined,
            undefined,
            warpObj
          );
          var rowSpan = getTextByPathList(tcNodes[j], ["attrs", "rowSpan"]);
          var colSpan = getTextByPathList(tcNodes[j], ["attrs", "gridSpan"]);
          var vMerge = getTextByPathList(tcNodes[j], ["attrs", "vMerge"]);
          var hMerge = getTextByPathList(tcNodes[j], ["attrs", "hMerge"]);
          if (rowSpan !== undefined) {
            tableHtml +=
              "<td rowspan='" + parseInt(rowSpan) + "'>" + text + "</td>";
          } else if (colSpan !== undefined) {
            tableHtml +=
              "<td colspan='" + parseInt(colSpan) + "'>" + text + "</td>";
          } else if (vMerge === undefined && hMerge === undefined) {
            tableHtml += "<td>" + text + "</td>";
          }
        }
      } else {
        var text = genTextBody(tcNodes);
        tableHtml += "<td>" + text + "</td>";
      }
      tableHtml += "</tr>";
    }
  } else {
    tableHtml += "<tr>";
    var tcNodes = trNodes["a:tc"];
    if (tcNodes.constructor === Array) {
      for (var j = 0; j < tcNodes.length; j++) {
        var text = genTextBody(tcNodes[j]);
        tableHtml += "<td>" + text + "</td>";
      }
    } else {
      var text = genTextBody(tcNodes);
      tableHtml += "<td>" + text + "</td>";
    }
    tableHtml += "</tr>";
  }

  return tableHtml;
}

async function genChart(node, warpObj, nodeType) {
  var order = getOrder(node, nodeType);
  var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
  var result =
    "<div id='chart" +
    chartID +
    "' class='block content' style='" +
    getPosition(xfrmNode, undefined, undefined) +
    getSize(xfrmNode, undefined, undefined) +
    " z-index: " +
    order +
    ";'></div>";

  var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
  var refName = warpObj["slideResObj"][rid]["target"];
  var content = await readXmlFile(warpObj["zip"], refName);
  var plotArea = getTextByPathList(content, [
    "c:chartSpace",
    "c:chart",
    "c:plotArea",
  ]);

  var chartData = null;
  for (var key in plotArea) {
    switch (key) {
      case "c:lineChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "lineChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:barChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "barChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:pieChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "pieChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:pie3DChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "pie3DChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:areaChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "areaChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:scatterChart":
        chartData = {
          type: "createChart",
          data: {
            chartID: "chart" + chartID,
            chartType: "scatterChart",
            chartData: extractChartData(plotArea[key]["c:ser"]),
          },
        };
        break;
      case "c:catAx":
        break;
      case "c:valAx":
        break;
      default:
    }
  }

  if (chartData !== null) {
    MsgQueue.push(chartData);
  }

  chartID++;
  return result;
}

function genDiagram(node, warpObj, nodeType) {
  var order = getOrder(node, nodeType);
  var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
  return (
    "<div class='block content' style='border: 1px dotted;" +
    getPosition(xfrmNode, undefined, undefined) +
    getSize(xfrmNode, undefined, undefined) +
    "'>TODO: diagram</div>"
  );
}

function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  var off = undefined;
  var x = -1,
    y = -1;

  if (slideSpNode !== undefined) {
    off = slideSpNode["a:off"]["attrs"];
  } else if (slideLayoutSpNode !== undefined) {
    off = slideLayoutSpNode["a:off"]["attrs"];
  } else if (slideMasterSpNode !== undefined) {
    off = slideMasterSpNode["a:off"]["attrs"];
  }

  if (off === undefined) {
    return "";
  } else {
    x = (parseInt(off["x"]) * 96) / 914400;
    y = (parseInt(off["y"]) * 96) / 914400;
    return isNaN(x) || isNaN(y) ? "" : "top:" + y + "px; left:" + x + "px;";
  }
}

function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  var ext = undefined;
  var w = -1,
    h = -1;
  var offset_padding_error = 0.5;
  if (slideSpNode !== undefined) {
    ext = slideSpNode["a:ext"]["attrs"];
  } else if (slideLayoutSpNode !== undefined) {
    ext = slideLayoutSpNode["a:ext"]["attrs"];
  } else if (slideMasterSpNode !== undefined) {
    ext = slideMasterSpNode["a:ext"]["attrs"];
  }

  if (ext === undefined) {
    return "";
  } else {
    w =
      (parseInt(ext["cx"]) * 96) / 914400 == 0
        ? 1
        : (parseInt(ext["cx"]) * 96) / 914400 + offset_padding_error;
    h =
      (parseInt(ext["cy"]) * 96) / 914400 == 0
        ? 1
        : (parseInt(ext["cy"]) * 96) / 914400 + offset_padding_error;
    return isNaN(w) || isNaN(h) ? "" : "width:" + w + "px; height:" + h + "px;";
  }
}

function getHorizontalAlign(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideMasterTextStyles
) {
  var algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
  if (algn === undefined) {
    var lvl = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:p",
      "a:pPr",
      "attrs",
      "lvl",
    ]);
    if (lvl !== undefined) {
      algn = getTextByPathList(slideLayoutSpNode, [
        "p:txBody",
        "a:lstStyle",
        "a:lvl" + (parseInt(lvl) + 1) + "pPr",
        "attrs",
        "algn",
      ]);
    } else {
      algn = getTextByPathList(slideLayoutSpNode, [
        "p:txBody",
        "a:lstStyle",
        "a:lvl1pPr",
        "attrs",
        "algn",
      ]);
    }
    if (algn === undefined) {
      algn = getTextByPathList(slideMasterSpNode, [
        "p:txBody",
        "a:p",
        "a:pPr",
        "attrs",
        "algn",
      ]);
      if (algn === undefined) {
        switch (type) {
          case "title":
          case "subTitle":
          case "ctrTitle":
            algn = getTextByPathList(slideMasterTextStyles, [
              "p:titleStyle",
              "a:lvl1pPr",
              "attrs",
              "alng",
            ]);
            break;
          default:
            algn = getTextByPathList(slideMasterTextStyles, [
              "p:otherStyle",
              "a:lvl1pPr",
              "attrs",
              "alng",
            ]);
        }
      }
    }
  }
  return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
}

function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {
  var anchor = getTextByPathList(node, [
    "p:txBody",
    "a:bodyPr",
    "attrs",
    "anchor",
  ]);
  if (anchor === undefined) {
    anchor = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:bodyPr",
      "attrs",
      "anchor",
    ]);
    if (anchor === undefined) {
      anchor = getTextByPathList(slideMasterSpNode, [
        "p:txBody",
        "a:bodyPr",
        "attrs",
        "anchor",
      ]);
      if (anchor === undefined && type === "ctrTitle") {
        anchor = "ctr";
      }
    }
  }
  return anchor === "ctr" ? "v-mid" : anchor === "b" ? "v-down" : "v-up";
}

function getPadding(node, slideLayoutSpNode, slideMasterSpNode, type) {
  var p_types = ["tIns", "rIns", "bIns", "lIns"];
  var result = "padding: ";
  for (let i = 0; i < p_types.length; i++) {
    if (i == 0 || i == 2) {
      var side = getTextByPathList(node, [
        "p:txBody",
        "a:bodyPr",
        "attrs",
        p_types[i],
      ]);
      if (isNaN(side) || side === undefined) {
        side = getTextByPathList(slideLayoutSpNode, [
          "p:txBody",
          "a:bodyPr",
          "attrs",
          p_types[i],
        ]);
        if (isNaN(side) || side === undefined) {
          side = getTextByPathList(slideMasterSpNode, [
            "p:txBody",
            "a:bodyPr",
            "attrs",
            p_types[i],
          ]);
          if (isNaN(side) || side === undefined) {
            side = "45720";
          }
        }
      }
    } else {
      var side = getTextByPathList(node, [
        "p:txBody",
        "a:bodyPr",
        "attrs",
        p_types[i],
      ]);
      if (isNaN(side) || side === undefined) {
        side = getTextByPathList(slideLayoutSpNode, [
          "p:txBody",
          "a:bodyPr",
          "attrs",
          p_types[i],
        ]);
        if (isNaN(side) || side === undefined) {
          side = getTextByPathList(slideMasterSpNode, [
            "p:txBody",
            "a:bodyPr",
            "attrs",
            p_types[i],
          ]);
          if (isNaN(side) || side === undefined) {
            side = "91440";
          }
        }
      }
    }
    result += (parseInt(side) * 96) / 914400 + "px ";
  }
  return result + ";box-sizing: border-box;";
}

function evaluateCommandText(text, tag) {
  var custom_command = tag;
  var result = `<div class="plugin-wrapper">`;
  var colors = [getSchemeColorFromTheme("a:tx1"), getSchemeColorFromTheme("a:bg1")];
  if (custom_command != undefined) {
    if (custom_command.includes("livecode")) {
      const lc = new LiveCode(custom_command,text,colors);
      result+= lc.create();
    }else if(custom_command.includes("website")){
      var website = custom_command.match(/&quot;(.*?)&quot;/g);
      website[0] = website[0].replace(/&quot;/g, "");
      result += `<span class="plugin-name">Live website - ${website[0]}</span>`;
      result+=`<iframe src="${website}" width="100%" height="100%"></iframe>`
    }else if (custom_command.includes("liveviz")) {
      const lv = new LiveViz(custom_command,text,colors,data_files_data,data_files_names);
      result+= lv.create();
      log_info += result + "\n";
    }else if(custom_command.includes("livequestion")){
      const lq = new LiveQuestion(custom_command,text,colors);
      result+= lq.create();
    }
    else {
      log_info +=
        `COMMAND ${custom_command} NOT IMPLEMENTED OR NOT RECOGNIZED` + "\n";
    }
  }
  return result + `</div><style>
    .plugin-wrapper{
      width:100%;
      height: 100%;
      outline: solid #${getSchemeColorFromTheme("a:tx1") != undefined ? getSchemeColorFromTheme("a:tx1"): "000"} 5px;
      outline-offset: 20px;
    }

    .plugin-name{
      position: absolute;
      left: -14.5px;
      top: -44px;
      padding: 7px;
      color: #${getSchemeColorFromTheme("a:bg1") != undefined ? getSchemeColorFromTheme("a:bg1"): "fff"};
      font-size: 14px;
      font-weight: 700;
      background: #${getSchemeColorFromTheme("a:tx1") != undefined ? getSchemeColorFromTheme("a:tx1"): "000"}; 
    }
  </style>`;
}

function getTextOrientation(node, slideLayoutSpNode, slideMasterSpNode) {
  var vert = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "vert"]);
  if (vert === undefined) {
    vert = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:bodyPr",
      "attrs",
      "vert",
    ]);
    if (vert === undefined) {
      vert = getTextByPathList(slideMasterSpNode, [
        "p:txBody",
        "a:bodyPr",
        "attrs",
        "vert",
      ]);
    }
  }
  if (vert !== undefined) {
    if (vert === "vert") {
      return "writing-mode: vertical-rl;";
    } else if (vert === "vert270") {
      return "writing-mode: vertical-rl;transform: rotate(180deg);margin-top: -6px;";
    } else {
      return "";
    }
  } else {
    return "writing-mode:initial;";
  }
}

function getFontType(node, type, slideMasterTextStyles) {
  var typeface = getTextByPathList(node, [
    "a:rPr",
    "a:latin",
    "attrs",
    "typeface",
  ]);

  if (typeface === undefined) {
    var fontSchemeNode = getTextByPathList(currentthemeContent, [
      "a:theme",
      "a:themeElements",
      "a:fontScheme",
    ]);
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      typeface = getTextByPathList(fontSchemeNode, [
        "a:majorFont",
        "a:latin",
        "attrs",
        "typeface",
      ]);
    } else if (type == "body") {
      typeface = getTextByPathList(fontSchemeNode, [
        "a:minorFont",
        "a:latin",
        "attrs",
        "typeface",
      ]);
    } else {
      typeface = getTextByPathList(fontSchemeNode, [
        "a:minorFont",
        "a:latin",
        "attrs",
        "typeface",
      ]);
    }
  }

  return typeface === undefined ? "inherit" : typeface;
}

function getFontColor(
  node,
  id,
  type,
  slideLayoutSpNode,
  slt,
  slideMasterTextStyles,
  lvl
) {
  var color = getRGBColorsFromNode(
    getTextByPathList(node, ["a:rPr", "a:solidFill"])
  );
  if (color[0] === undefined) {
    var color_node = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "a:solidFill",
    ]);
    color = getRGBColorsFromNode(color_node);
    if (color[0] === undefined) {
      var color_node = getTextByPathList(slt["idTable"][id], [
        "p:style",
        "a:fontRef",
      ]);
      color = getRGBColorsFromNode(color_node);
      if (color[0] === undefined) {
        color = getRGBColorsFromNode(
          getTextByPathList(slt["typeTable"][type], [
            "p:txBody",
            "a:lstStyle",
            lvl,
            "a:defRPr",
            "a:solidFill",
          ])
        );
        if (color[0] === undefined) {
          if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            color = getRGBColorsFromNode(
              getTextByPathList(slideMasterTextStyles, [
                "p:titleStyle",
                lvl,
                "a:defRPr",
                "a:solidFill",
              ])
            );
          } else if (type == "body") {
            color = getRGBColorsFromNode(
              getTextByPathList(slideMasterTextStyles, [
                "p:bodyStyle",
                lvl,
                "a:defRPr",
                "a:solidFill",
              ])
            );
          } else if (type === undefined) {
            color = getRGBColorsFromNode(
              getTextByPathList(slideMasterTextStyles, [
                "p:otherStyle",
                lvl,
                "a:defRPr",
                "a:solidFill",
              ])
            );
          }
        }
      }
    }
  }
  return color[0] === undefined ? "#000" : "#" + color[0].toHex();
}

function getFontSize(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var fontSize = undefined;
  if (node["a:rPr"] !== undefined) {
    fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
  }

  if (isNaN(fontSize) || fontSize === undefined) {
    var sz = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "sz",
    ]);
    fontSize = parseInt(sz) / 100;
  }

  if (isNaN(fontSize) || fontSize === undefined) {
    var sz = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "sz",
    ]);
    fontSize = parseInt(sz) / 100;
  }

  if (isNaN(fontSize) || fontSize === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "sz",
      ]);
    } else if (type == "body") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "sz",
      ]);
    } else if (type == "dt" || type == "sldNum") {
      var sz = "1200";
    } else if (type === undefined) {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "sz",
      ]);
    }
    fontSize = parseInt(sz) / 100;
  }

  var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
  if (baseline !== undefined && !isNaN(fontSize)) {
    fontSize -= 10;
  }

  return isNaN(fontSize) ? "inherit" : fontSize + "pt";
}
function getMarginsSize(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var marL = undefined;
  if (node["a:pPr"] !== undefined) {
    marL = (parseInt(node["a:pPr"]["attrs"]["marL"]) * 96) / 914400;
  }

  if (isNaN(marL) || marL === undefined) {
    var sz = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "attrs",
      "marL",
    ]);
    marL = (parseInt(sz) * 96) / 914400;
  }

  if (isNaN(marL) || marL === undefined) {
    var sz = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "attrs",
      "marL",
    ]);
    marL = (parseInt(sz) * 96) / 914400;
  }

  if (isNaN(marL) || marL === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "attrs",
        "marL",
      ]);
    } else if (type == "body") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "attrs",
        "marL",
      ]);
    } else if (type == "dt" || type == "sldNum") {
      var sz = "1200";
    } else if (type === undefined) {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "attrs",
        "marL",
      ]);
    }
    marL = (parseInt(sz) * 96) / 914400;
  }
  if (isNaN(marL) || marL === undefined) {
    marL = (347663 * 96) / 914400;
  }

  return "0 0 0 " + marL + "px";
}

function getIndentSize(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var indent = undefined;
  if (node["a:pPr"] !== undefined) {
    indent = (parseInt(node["a:pPr"]["attrs"]["indent"]) * 96) / 914400;
  }

  if (isNaN(indent) || indent === undefined) {
    var sz = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "attrs",
      "indent",
    ]);
    indent = (parseInt(sz) * 96) / 914400;
  }

  if (isNaN(indent) || indent === undefined) {
    var sz = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "attrs",
      "indent",
    ]);
    indent = (parseInt(sz) * 96) / 914400;
  }

  if (isNaN(indent) || indent === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "attrs",
        "indent",
      ]);
    } else if (type == "body") {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "attrs",
        "indent",
      ]);
    } else if (type == "dt" || type == "sldNum") {
      var sz = "1200";
    } else if (type === undefined) {
      var sz = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "attrs",
        "indent",
      ]);
    }
    indent = (parseInt(sz) * 96) / 914400;
  }

  return indent + "px";
}

function getTextCap(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var caps = undefined;
  if (node["a:rPr"] !== undefined) {
    caps = node["a:rPr"]["attrs"]["cap"];
  }

  if (caps === undefined) {
    caps = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "cap",
    ]);
  }

  if (caps === undefined) {
    caps = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "cap",
    ]);
  }

  if (caps === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      caps = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "cap",
      ]);
    } else if (type == "body") {
      caps = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "cap",
      ]);
    } else if (type === undefined) {
      caps = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "cap",
      ]);
    }
  }
  return caps === "all" ? "uppercase" : caps === "small" ? "lowercase" : "none";
}

function getFontBold(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var b = undefined;
  if (node["a:rPr"] !== undefined) {
    b = node["a:rPr"]["attrs"]["b"];
  }

  if (b === undefined) {
    b = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "b",
    ]);
  }

  if (b === undefined) {
    b = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "b",
    ]);
  }

  if (b === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      b = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "b",
      ]);
    } else if (type == "body") {
      b = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "b",
      ]);
    } else if (type === undefined) {
      b = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "b",
      ]);
    }
  }
  return b === "1" ? "bold" : "normal";
}

function getFontItalic(
  node,
  slideLayoutSpNode,
  slideMasterSpNode,
  type,
  slideLayoutTables,
  slideMasterTextStyles,
  lvl
) {
  var ita = undefined;
  if (node["a:rPr"] !== undefined) {
    ita = node["a:rPr"]["attrs"]["i"];
  }

  if (ita === undefined) {
    ita = getTextByPathList(slideLayoutSpNode, [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "i",
    ]);
  }

  if (ita === undefined) {
    ita = getTextByPathList(slideLayoutTables["typeTable"][type], [
      "p:txBody",
      "a:lstStyle",
      lvl,
      "a:defRPr",
      "attrs",
      "i",
    ]);
  }

  if (ita === undefined) {
    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
      ita = getTextByPathList(slideMasterTextStyles, [
        "p:titleStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "i",
      ]);
    } else if (type == "body") {
      ita = getTextByPathList(slideMasterTextStyles, [
        "p:bodyStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "i",
      ]);
    } else if (type === undefined) {
      ita = getTextByPathList(slideMasterTextStyles, [
        "p:otherStyle",
        lvl,
        "a:defRPr",
        "attrs",
        "i",
      ]);
    }
  }
  return ita === "1" ? "italic" : "normal";
}

function getLineSpacing(node, type, slideMasterTextStyles) {
  var linespacing = undefined;
  linespacing = getTextByPathList(node, [
    "a:pPr",
    "a:lnSpc",
    "a:spcPct",
    "attrs",
    "val",
  ]);
  if (linespacing === undefined) {
    switch (type) {
      case "title":
      case "subTitle":
      case "ctrTitle":
        linespacing = getTextByPathList(slideMasterTextStyles, [
          "p:titleStyle",
          "a:lvl1pPr",
          "a:lnSpc",
          "a:spcPct",
          "attrs",
          "val",
        ]);
        break;
      case "body":
        linespacing = getTextByPathList(slideMasterTextStyles, [
          "p:bodyStyle",
          "a:lvl1pPr",
          "a:lnSpc",
          "a:spcPct",
          "attrs",
          "val",
        ]);
        break;
    }
  }

  if (linespacing != undefined) {
    return parseInt(linespacing) / 2000 + "px";
  } else {
    return "initial";
  }
}

function getFontDecoration(node, type, slideMasterTextStyles) {
  return node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng"
    ? "underline"
    : "initial";
}

function getTextVerticalAlign(node, type, slideMasterTextStyles) {
  var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
  return baseline === undefined ? "baseline" : parseInt(baseline) / 1000 + "%";
}

function getBorder(node, slideLayoutSpNode, slideMasterSpNode, isSvgMode) {
  var cssText = "border: ";
  var selectNode = undefined;
  // Border width: 1pt = 12700, default = 0.75pt
  var borderWidth =
    parseInt(getTextByPathList(node, ["p:spPr", "a:ln", "attrs", "w"])) / 12700;
  if (isNaN(borderWidth) || borderWidth === undefined) {
    borderWidth =
      parseInt(
        getTextByPathList(slideLayoutSpNode, ["p:spPr", "a:ln", "attrs", "w"])
      ) / 12700;
    if (isNaN(borderWidth) || borderWidth === undefined) {
      borderWidth =
        parseInt(
          getTextByPathList(slideMasterSpNode, ["p:spPr", "a:ln", "attrs", "w"])
        ) / 12700;
      if (isNaN(borderWidth) || borderWidth === undefined) {
        return "border: none;";
      } else {
        selectNode = slideMasterSpNode;
      }
    } else {
      selectNode = slideLayoutSpNode;
    }
  } else {
    selectNode = node;
  }
  if (isNaN(borderWidth) || borderWidth < 1) {
    cssText += "0pt ";
  } else {
    cssText += borderWidth + "pt ";
  }

  // Border color
  var borderColor = getTextByPathList(selectNode, [
    "p:spPr",
    "a:ln",
    "a:solidFill",
    "a:srgbClr",
    "attrs",
    "val",
  ]);
  if (borderColor === undefined) {
    var schemeClrNode = getTextByPathList(selectNode, [
      "p:spPr",
      "a:ln",
      "a:solidFill",
      "a:schemeClr",
    ]);
    var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);
    var borderColor = getSchemeColorFromTheme(schemeClr);
  }

  // 2. drawingML namespace
  if (borderColor === undefined) {
    var schemeClrNode = getTextByPathList(selectNode, [
      "p:style",
      "a:lnRef",
      "a:schemeClr",
    ]);
    var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);
    var borderColor = getSchemeColorFromTheme(schemeClr);

    if (borderColor !== undefined) {
      var shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
      if (shade !== undefined) {
        shade = parseInt(shade) / 100000;
        var color = new colz.Color("#" + borderColor);
        color.setLum(color.hsl.l * shade);
        borderColor = color.hex.replace("#", "");
      }
    }
  }

  if (borderColor === undefined) {
    if (isSvgMode) {
      borderColor = "none";
    } else {
      borderColor = "transparent";
    }
  } else {
    borderColor = "#" + borderColor;
  }
  cssText += " " + borderColor + " ";

  // Border type
  var borderType = getTextByPathList(selectNode, [
    "p:spPr",
    "a:ln",
    "a:prstDash",
    "attrs",
    "val",
  ]);
  var strokeDasharray = "0";
  switch (borderType) {
    case "solid":
      cssText += "solid";
      strokeDasharray = "0";
      break;
    case "dash":
      cssText += "dashed";
      strokeDasharray = "5";
      break;
    case "dashDot":
      cssText += "dashed";
      strokeDasharray = "5, 5, 1, 5";
      break;
    case "dot":
      cssText += "dotted";
      strokeDasharray = "1, 5";
      break;
    case "lgDash":
      cssText += "dashed";
      strokeDasharray = "10, 5";
      break;
    case "lgDashDotDot":
      cssText += "dashed";
      strokeDasharray = "10, 5, 1, 5, 1, 5";
      break;
    case "sysDash":
      cssText += "dashed";
      strokeDasharray = "5, 2";
      break;
    case "sysDashDot":
      cssText += "dashed";
      strokeDasharray = "5, 2, 1, 5";
      break;
    case "sysDashDotDot":
      cssText += "dashed";
      strokeDasharray = "5, 2, 1, 5, 1, 5";
      break;
    case "sysDot":
      cssText += "dotted";
      strokeDasharray = "2, 5";
      break;
    default:
      cssText += "solid";
      strokeDasharray = "0";
  }

  if (isSvgMode) {
    return {
      color: borderColor,
      width: borderWidth,
      type: borderType,
      strokeDasharray: strokeDasharray,
    };
  } else {
    return cssText + ";";
  }
}

function getSlideBackgroundFill(
  slideContent,
  slideLayoutContent,
  slideMasterContent
) {
  var bgColor = getSolidFill(
    getTextByPathList(slideContent, [
      "p:sld",
      "p:cSld",
      "p:bg",
      "p:bgPr",
      "a:solidFill",
    ])
  );
  if (bgColor === undefined) {
    bgColor = getSolidFill(
      getTextByPathList(slideLayoutContent, [
        "p:sldLayout",
        "p:cSld",
        "p:bg",
        "p:bgPr",
        "a:solidFill",
      ])
    );
    if (bgColor === undefined) {
      bgColor = getSolidFill(
        getTextByPathList(slideMasterContent, [
          "p:sldMaster",
          "p:cSld",
          "p:bg",
          "p:bgPr",
          "a:solidFill",
        ])
      );
      if (bgColor === undefined) {
        bgColor = "FFF";
      }
    }
  }
  return bgColor;
}

function getShapeFill(node, slideLayoutSpNode, slideMasterSpNode, isSvgMode) {
  // 1. presentationML
  // p:spPr [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
  // From slide
  if (getTextByPathList(node, ["p:spPr", "a:noFill"]) !== undefined) {
    return isSvgMode ? "none" : "background-color: initial;";
  }
  var fillColor = undefined;
  if (fillColor === undefined) {
    fillColor = getRGBColorsFromNode(
      getTextByPathList(node, ["p:spPr", "a:solidFill"])
    )[0];
  }

  if (fillColor === undefined) {
    fillColor = getRGBColorsFromNode(
      getTextByPathList(slideLayoutSpNode, ["p:spPr", "a:solidFill"])
    )[0];
  }

  if (fillColor === undefined) {
    fillColor = getRGBColorsFromNode(
      getTextByPathList(slideMasterSpNode, ["p:spPr", "a:solidFill"])
    )[0];
  }

  // // From theme
  // if (fillColor === undefined) {
  //     var schemeClr = "a:" + getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "attrs", "val"]);
  //     fillColor = getSchemeColorFromTheme(schemeClr);
  // }

  // // 2. drawingML namespace
  // if (fillColor === undefined) {
  //     var schemeClr = "a:" + getTextByPathList(node, ["p:style", "a:fillRef", "a:schemeClr", "attrs", "val"]);
  //     fillColor = getSchemeColorFromTheme(schemeClr);
  // }
  if (fillColor !== undefined) {
    // fillColor = "#" + fillColor;

    // // Apply shade or tint
    // // TODO: 較淺, 較深 80%
    // var lumMod = parseInt(getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumMod", "attrs", "val"]));
    // var lumOff = parseInt(getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumOff", "attrs", "val"]));
    // var alpha = parseInt(getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:alpha", "attrs", "val"]));
    // fillColor = applyLumModify(fillColor, lumMod, lumOff, alpha);
    if (isSvgMode) {
      return fillColor;
    } else {
      return "background-color: " + fillColor + ";";
    }
  } else {
    if (isSvgMode) {
      return "none";
    } else {
      return "background-color: transparent;";
    }
  }
}

function getOrder(node, nodeType) {
  var order = node["attrs"]["order"];
  if (nodeType == "master" || nodeType == "layout") {
    if (order > biggestMasterLayoutOrder) {
      biggestMasterLayoutOrder = order;
      return order;
    }
  } else {
    return biggestMasterLayoutOrder + order;
  }
}

function getSolidFill(solidFill) {
  if (solidFill === undefined) {
    return undefined;
  }

  var color = "FFF";

  if (solidFill["a:srgbClr"] !== undefined) {
    color = getTextByPathList(solidFill["a:srgbClr"], ["attrs", "val"]);
  } else if (solidFill["a:schemeClr"] !== undefined) {
    var schemeClr =
      "a:" + getTextByPathList(solidFill["a:schemeClr"], ["attrs", "val"]);
    color = getSchemeColorFromTheme(schemeClr);
  }

  return color;
}

function getSchemeColorFromTheme(schemeClr) {
  // TODO: <p:clrMap ...> in slide master
  // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1"
  switch (schemeClr) {
    case "a:tx1":
      schemeClr = "a:dk1";
      break;
    case "a:tx2":
      schemeClr = "a:dk2";
      break;
    case "a:bg1":
      schemeClr = "a:lt1";
      break;
    case "a:bg2":
      schemeClr = "a:lt2";
      break;
  }
  var refNode = getTextByPathList(currentthemeContent, [
    "a:theme",
    "a:themeElements",
    "a:clrScheme",
    schemeClr,
  ]);
  var color = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
  if (color === undefined) {
    color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
  }
  return color;
}

function extractChartData(serNode) {
  var dataMat = new Array();

  if (serNode === undefined) {
    return dataMat;
  }

  if (serNode["c:xVal"] !== undefined) {
    var dataRow = new Array();
    eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (
      innerNode,
      index
    ) {
      dataRow.push(parseFloat(innerNode["c:v"]));
      return "";
    });
    dataMat.push(dataRow);
    dataRow = new Array();
    eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (
      innerNode,
      index
    ) {
      dataRow.push(parseFloat(innerNode["c:v"]));
      return "";
    });
    dataMat.push(dataRow);
  } else {
    eachElement(serNode, function (innerNode, index) {
      var dataRow = new Array();
      var colName =
        getTextByPathList(innerNode, [
          "c:tx",
          "c:strRef",
          "c:strCache",
          "c:pt",
          "c:v",
        ]) || index;

      // Category (string or number)
      var rowNames = {};
      if (
        getTextByPathList(innerNode, [
          "c:cat",
          "c:strRef",
          "c:strCache",
          "c:pt",
        ]) !== undefined
      ) {
        eachElement(
          innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"],
          function (innerNode, index) {
            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
            return "";
          }
        );
      } else if (
        getTextByPathList(innerNode, [
          "c:cat",
          "c:numRef",
          "c:numCache",
          "c:pt",
        ]) !== undefined
      ) {
        eachElement(
          innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"],
          function (innerNode, index) {
            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
            return "";
          }
        );
      }

      // Value
      if (
        getTextByPathList(innerNode, [
          "c:val",
          "c:numRef",
          "c:numCache",
          "c:pt",
        ]) !== undefined
      ) {
        eachElement(
          innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"],
          function (innerNode, index) {
            dataRow.push({
              x: innerNode["attrs"]["idx"],
              y: parseFloat(innerNode["c:v"]),
            });
            return "";
          }
        );
      }

      dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
      return "";
    });
  }

  return dataMat;
}

function getTextFromNode(node) {
  if (node === undefined || node["a:t"] === undefined) {
    return "";
  }
  var text = node["a:t"];
  if (typeof text !== "string") {
    text = getTextByPathList(node, ["a:fld", "a:t"]);
    if (typeof text !== "string") {
      text = "";
    }
  }
  return text;
}

/**
 * This funcion applies a function to an effect defined here http://officeopenxml.com/drwPic-effects.php
 * @param {node} node
 * @param {Jimp} image 
 */

function applyImageEffects(node, image) {
  var effectsNode = node["p:blipFill"]["a:blip"];

  if (effectsNode != undefined) {
    for (var key in effectsNode) {
      if (key == "attrs" || key == "a:extLst") {
        continue;
      }
      switch (key) {
        case "a:alphaInv":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:alphaRepl":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:alphaModFix":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:biLevel":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:blur":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:clrChange":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:clrRepl":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:duotone":
          var colors = getRGBColorsFromNode(effectsNode[key]);
          duotone(image, colors[0], colors[1]);
          break;
        case "a:grayscl":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:lum":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        case "a:tint":
          log_info += `ERROR: ${key} Not yet implemented!` + "\n";
          break;
        default:
          log_info += `ERROR: Effect type ${key} not defined!` + "\n";
      }
    }
  }
}

// ===== Node functions =====
/**
 * getTextByPathStr
 * @param {Object} node
 * @param {string} pathStr
 */
function getTextByPathStr(node, pathStr) {
  return getTextByPathList(node, pathStr.trim().split(/\s+/));
}

/**
 * getTextByPathList
 * @param {Object} node
 * @param {string Array} path
 */
function getTextByPathList(node, path) {
  if (path.constructor !== Array) {
    throw Error("Error of path type! path is not array.");
  }

  if (node === undefined) {
    return undefined;
  }

  var l = path.length;
  for (var i = 0; i < l; i++) {
    node = node[path[i]];
    if (node === undefined) {
      return undefined;
    }
  }

  return node;
}

/**
 * eachElement
 * @param {Object} node
 * @param {function} doFunction
 */
function eachElement(node, doFunction) {
  if (node === undefined) {
    return;
  }
  var result = "";
  if (node.constructor === Array) {
    var l = node.length;
    for (var i = 0; i < l; i++) {
      result += doFunction(node[i], i);
    }
  } else {
    result += doFunction(node, 0);
  }
  return result;
}

// ===== Color functions =====
/**
 * applyShade effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} shadeValue
 */
function applyShade(color, shadeValue) {
  color._r = color._r * (shadeValue / 100000);
  color._g = color._g * (shadeValue / 100000);
  color._b = color._b * (shadeValue / 100000);
  return color;
}

/**
 * applyTint effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} tintValue
 */
function applyTint(color, tintValue) {
  color._r = color._r * (tintValue / 100000) + (1 - tintValue / 100000);
  color._g = color._g * (tintValue / 100000) + (1 - tintValue / 100000);
  color._b = color._b * (tintValue / 100000) + (1 - tintValue / 100000);
  return color;
}

/**
 * applySaturation effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} satValue
 */
function applySaturation(color, satValue) {
  color = color.toHsv();
  color.s = color.s * (satValue / 100000);
  color = tinycolor(color);
  return color;
}

/**
 * applyAlpha effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} alphaValue
 */
function applyAlpha(color, alphaValue) {
  color.setAlpha(alphaValue);
  color = tinycolor(color);
  return color;
}

/**
 * applyLumMod effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} lumModValue
 */
function applyLumMod(color, lumModValue) {
  color = color.toHsl();
  color.l = color.l * (lumModValue / 100000);
  color = tinycolor(color);
  return color;
}

/**
 * applyLumOff effect. See more in http://officeopenxml.com/drwSp-effects.php
 * @param {Color} color
 * @param {number} lumOffValue
 */
function applyLumOff(color, lumOffValue) {
  color = color.toHsl();
  color.l = color.l + lumOffValue / 100000;
  color = tinycolor(color);
  return color;
}

/**
 * This funcion applies a function to a color effect. http://officeopenxml.com/drwOverview.php
 * @param {node} node
 * @param {tinycolor} color 
 */
function applyColorEffects(node, color) {
  for (var key in node) {
    if (key == "attrs") {
      continue;
    }
    switch (key) {
      case "a:alpha":
        color = applyAlpha(color, node[key]["attrs"]["val"]);
        break;
      case "a:lumMod":
        color = applyLumMod(color, node[key]["attrs"]["val"]);
        break;
      case "a:lumOff":
        color = applyLumOff(color, node[key]["attrs"]["val"]);
        break;
      case "a:satMod":
        color = applySaturation(color, node[key]["attrs"]["val"]);
        break;
      case "a:shade":
        color = applyShade(color, node[key]["attrs"]["val"]);
        break;
      case "a:tint":
        color = applyTint(color, node[key]["attrs"]["val"]);
        break;
      default:
        log_info += `ERROR: Color effect key -${key}- not defined!` + "\n";
    }
  }
  return color;
}

/**
 * This funcion make a gradient color effect.
 * @param {tinycolor} tone1
 * @param {tinycolor} tone2 
 */
function gradientMap(tone1, tone2) {
  var gradient = [];
  for (var i = 0; i < 256 * 4; i += 4) {
    gradient[i] = Math.round(
      ((256 - i / 4) * tone1._r + (i / 4) * tone2._r) / 256
    );
    gradient[i + 1] = Math.round(
      ((256 - i / 4) * tone1._g + (i / 4) * tone2._g) / 256
    );
    gradient[i + 2] = Math.round(
      ((256 - i / 4) * tone1._b + (i / 4) * tone2._b) / 256
    );
    gradient[i + 3] = 255;
  }
  return gradient;
}

/**
 * This funcion make a gradient color effect. http://officeopenxml.com/drwPic-effects.php
 * @param {jimp} img
 * @param {tinycolor} tone1
 * @param {tinycolor} tone2 
 */

function duotone(img, tone1, tone2) {
  img.greyscale();
  var gradient = gradientMap(tone1, tone2);
  img.scan(0, 0, img.bitmap.width, img.bitmap.height, function (x, y, idx) {
    this.bitmap.data[idx + 0] = gradient[this.bitmap.data[idx + 0] * 4];
    this.bitmap.data[idx + 1] = gradient[this.bitmap.data[idx + 1] * 4 + 1];
    this.bitmap.data[idx + 2] = gradient[this.bitmap.data[idx + 2] * 4 + 2];
  });
  return img;
}

/**
 * This funcion gets the colors from a node . http://officeopenxml.com/drwOverview.php
 * @param {node} node
 */
function getRGBColorsFromNode(node) {
  var colors = [];
  for (var key in node) {
    var color = undefined;
    if (key == "attrs") {
      continue;
    }
    switch (key) {
      case "a:prstClr":
        var value = node[key]["attrs"]["val"];
        color = tinycolor(value);
        break;
      case "a:sysClr":
      case "a:schemeClr":
        var value = "a:" + node[key]["attrs"]["val"];
        color = tinycolor(getSchemeColorFromTheme(value));
        break;
      case "a:hslClr":
      case "a:scrgbClr":
      case "a:srgbClr":
        var value = node[key]["attrs"]["val"];
        color = tinycolor(value);
        break;
      default:
        log_info += `COLOR FORMAT ${key} NOT RECOGNIZED!` + "\n";
    }
    if (color != undefined) {
      color = applyColorEffects(node[key], color);
      colors.push(color);
    }
  }
  return colors;
} 

function extractFileExtension(filename) {
  return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
}

/**
 * This funcion processes the input from the command line and prepares an array of arguments for the command line version
 */
function processCommandLine(){
  process.stdout.write("\x1B[?25l");
  process.stdout.write("0% ");
  for (let i = 0; i < size; i++) {
    process.stdout.write("\u2591");
  }
  process.stdout.write(" 100%");
  rdl.cursorTo(process.stdout, cursor, 0);
  
  if (process.argv[2] != undefined && process.argv[2].slice(-4) == "pptx") {
    powerPointFile = process.argv[2];
  } else {
    console.log(
      "A pptx file was not provided or is not a pptx file! Please correct it on the command line."
    );
    process.exit(0);
  }
  
  var myArgs = process.argv.slice(2);
  
  if (myArgs.includes(DS_CONFIG.audience_interaction_command)){
    audience_interaction_plugin = true;
    myArgs = myArgs.filter(input => input !== DS_CONFIG.audience_interaction_command);
  }
  return myArgs;
}


function CSVArrayToJSON(array) {
  var result = [];
  for (let i = 1; i < array.length; i++) {
    var json_obj = {};
    for (let j = 0; j < array[0].length; j++) {
      json_obj[array[0][j]] = array[i][j];
    }
    result.push(json_obj);
  }
  return result;
}

function getFilePathsRecursively(dir){
  // if (isBrowser()) {
  //   throw new Error('getFilePathsRecursively is not supported in browser');
  // }

  // returns a flat array of absolute paths of all files recursively contained in the dir
  let results = [];
  let list = fs.readdirSync(dir);

  var pending = list.length;
  if (!pending) return results;

  for (let file of list) {
    file = path.resolve(dir, file);

    let stat = fs.lstatSync(file);

    if (stat && stat.isDirectory()) {
      results = results.concat(getFilePathsRecursively(file));
    } else {
      results.push(file);
    }

    if (!--pending) return results;
  }

  return results;
};

function getZipOfFolder(dir){
  // if (isBrowser()) {
  //   throw new Error('getZipOfFolder is not supported in browser');
  // }

  // returns a JSZip instance filled with contents of dir.

  let allPaths = getFilePathsRecursively(dir);
  let zip = new JSZip();

  for (let filePath of allPaths) {
    let addPath = path.relative(path.join(dir, '..'), filePath); // use this instead if you want the source folder itself in the zip
    // let addPath = path.relative(dir, filePath); // use this instead if you don't want the source folder itself in the zip
    let data = fs.readFileSync(filePath);
    let stat = fs.lstatSync(filePath);
    let permissions = stat.mode;

    if (stat.isSymbolicLink()) {
      zip.file(addPath, fs.readlinkSync(filePath), {
        unixPermissions: parseInt('120755', 8), // This permission can be more permissive than necessary for non-executables but we don't mind.
        dir: stat.isDirectory()
      });
    } else {
      zip.file(addPath, data, {
        unixPermissions: permissions,
        dir: stat.isDirectory()
      });
    }
  }

  return zip;
};

function insertFolderInZip(target_zip,dir,folder){
  // if (isBrowser()) {
  //   throw new Error('getZipOfFolder is not supported in browser');
  // }

  // returns a JSZip instance filled with contents of dir.

  let allPaths = getFilePathsRecursively(dir);
  let zip = target_zip.folder(folder);
  
  for (let filePath of allPaths) {
    let addPath = path.relative(path.join(dir, '..'), filePath); // use this instead if you want the source folder itself in the zip
    // let addPath = path.relative(dir, filePath); // use this instead if you don't want the source folder itself in the zip
    addPath = addPath.replace(/\\/g, "/");
    
    let data = fs.readFileSync(filePath);
    let stat = fs.lstatSync(filePath);
    let permissions = stat.mode;

    if (stat.isSymbolicLink()) {
      zip.file(addPath, fs.readlinkSync(filePath), {
        unixPermissions: parseInt('120755', 8), // This permission can be more permissive than necessary for non-executables but we don't mind.
        dir: stat.isDirectory()
      });
    } else {
      zip.file(addPath, data, {
        unixPermissions: permissions,
        dir: stat.isDirectory()
      });
    }
  }

  return zip;
};
module.exports = Dynaslides;