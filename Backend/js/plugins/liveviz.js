class LiveViz {
    constructor(command,text,colors,data_files_data,data_files_names) {
        this.command = command;
        this.text = text;
        this.colors = colors;
        this.data_files_data = data_files_data;
        this.data_files_names = data_files_names;
      }
    /**
    * This is where the html plugin funcionality is added.
    * @return
    *   The HTML in a text format
    */
    create(){
        var result = "";
        var data_filename = this.command.match(/&quot;(.*?)&quot;/g);
        data_filename[0] = data_filename[0].replace(/&quot;/g, "");
        var data =
        this.data_files_data[
            this.data_files_names.findIndex((a) => a.includes(data_filename[0]))
          ];
        result += `<span class="plugin-name">Live visualization - ${data_filename[0]}</span>`;
        result += `
              <div id="liveviz" style="height: 100%;">
              <div class="settings-liveviz">
                  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M24 13.616v-3.232c-1.651-.587-2.694-.752-3.219-2.019v-.001c-.527-1.271.1-2.134.847-3.707l-2.285-2.285c-1.561.742-2.433 1.375-3.707.847h-.001c-1.269-.526-1.435-1.576-2.019-3.219h-3.232c-.582 1.635-.749 2.692-2.019 3.219h-.001c-1.271.528-2.132-.098-3.707-.847l-2.285 2.285c.745 1.568 1.375 2.434.847 3.707-.527 1.271-1.584 1.438-3.219 2.02v3.232c1.632.58 2.692.749 3.219 2.019.53 1.282-.114 2.166-.847 3.707l2.285 2.286c1.562-.743 2.434-1.375 3.707-.847h.001c1.27.526 1.436 1.579 2.019 3.219h3.232c.582-1.636.75-2.69 2.027-3.222h.001c1.262-.524 2.12.101 3.698.851l2.285-2.286c-.744-1.563-1.375-2.433-.848-3.706.527-1.271 1.588-1.44 3.221-2.021zm-12 2.384c-2.209 0-4-1.791-4-4s1.791-4 4-4 4 1.791 4 4-1.791 4-4 4z"/>
                  </svg>
                  <div class="inner-settings-liveviz">
                      <div class="viz-types">
                          <svg version="1.1" id="Capa_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="24" height="24" x="0px" y="0px" viewBox="0 0 477.88 477.88" style="enable-background:new 0 0 477.88 477.88;" xml:space="preserve">
                              <g>
                                  <g>
                                      <path d="M472.897,124.269c-0.01-0.01-0.02-0.02-0.03-0.031l-0.017,0.017l-68.267-68.267c-6.78-6.548-17.584-6.36-24.132,0.42
                                          c-6.388,6.614-6.388,17.099,0,23.713l39.151,39.151h-95.334c-65.948,0.075-119.391,53.518-119.467,119.467
                                          c-0.056,47.105-38.228,85.277-85.333,85.333h-102.4C7.641,324.072,0,331.713,0,341.139s7.641,17.067,17.067,17.067h102.4
                                          c65.948-0.075,119.391-53.518,119.467-119.467c0.056-47.105,38.228-85.277,85.333-85.333h95.334l-39.134,39.134
                                          c-6.78,6.548-6.968,17.353-0.419,24.132c6.548,6.78,17.353,6.968,24.132,0.419c0.142-0.137,0.282-0.277,0.419-0.419l68.267-68.267
                                          C479.54,141.748,479.553,130.942,472.897,124.269z"/>
                                  </g>
                              </g>
                              <g>
                                  <g>
                                      <path d="M472.897,329.069c-0.01-0.01-0.02-0.02-0.03-0.03l-0.017,0.017l-68.267-68.267c-6.78-6.548-17.584-6.36-24.132,0.42
                                          c-6.388,6.614-6.388,17.099,0,23.712l39.151,39.151h-95.334c-20.996,0.015-41.258-7.721-56.9-21.726
                                          c-7.081-6.222-17.864-5.525-24.086,1.555c-6.14,6.988-5.553,17.605,1.319,23.874c21.898,19.614,50.269,30.451,79.667,30.43h95.334
                                          l-39.134,39.134c-6.78,6.548-6.968,17.352-0.42,24.132c6.548,6.78,17.352,6.968,24.132,0.42c0.142-0.138,0.282-0.277,0.42-0.42
                                          l68.267-68.267C479.54,346.548,479.553,335.742,472.897,329.069z"/>
                                  </g>
                              </g>
                              <g>
                                  <g>
                                      <path d="M199.134,149.702c-21.898-19.614-50.269-30.451-79.667-30.43h-102.4C7.641,119.272,0,126.913,0,136.339
                                          c0,9.426,7.641,17.067,17.067,17.067h102.4c20.996-0.015,41.258,7.721,56.9,21.726c7.081,6.222,17.864,5.525,24.086-1.555
                                          C206.593,166.588,206.006,155.971,199.134,149.702z"/>
                                  </g>
                              </g>
                          </svg>
                          <div class="inner-viz-types">
                              <span onclick="changeChartType(1)" style="margin-right:10px;width:25px; height:25px; display: inline-flex; justify-content:center; align-items:center; background-color: red; color:white;">1</span>
                              <span onclick="changeChartType(2)" style="margin-right:10px;width:25px; height:25px; display: inline-flex; justify-content:center; align-items:center; background-color: red; color:white;">2</span>
                              <span onclick="changeChartType(3)" style="margin-right:10px;width:25px; height:25px; display: inline-flex; justify-content:center; align-items:center; background-color: red; color:white;">3</span>
                          </div>
                      </div>
                      <div class="add-viz-data">
                          <svg height="24" viewBox="0 0 512 512" width="24" xmlns="http://www.w3.org/2000/svg"><path d="m256 0c-141.164062 0-256 114.835938-256 256s114.835938 256 256 256 256-114.835938 256-256-114.835938-256-256-256zm112 277.332031h-90.667969v90.667969c0 11.777344-9.554687 21.332031-21.332031 21.332031s-21.332031-9.554687-21.332031-21.332031v-90.667969h-90.667969c-11.777344 0-21.332031-9.554687-21.332031-21.332031s9.554687-21.332031 21.332031-21.332031h90.667969v-90.667969c0-11.777344 9.554687-21.332031 21.332031-21.332031s21.332031 9.554687 21.332031 21.332031v90.667969h90.667969c11.777344 0 21.332031 9.554687 21.332031 21.332031s-9.554687 21.332031-21.332031 21.332031zm0 0"/></svg>
                          <div class="add-viz-data-form"></div>
                      </div>
                  </div>
              </div>
          </div>
          <style>
              .settings-liveviz{
                position: absolute;
                top: -9px;
                right: -186px;
                cursor: pointer;
                text-align: left;      
              }
              .inner-settings-liveviz, .inner-viz-types, .add-viz-data-form{
                  visibility: hidden;
                  opacity: 0;
                  transition: visibility,opacity 0.3s linear
              }
              .inner-settings-liveviz, .add-viz-data{
                  margin-top: 25px;
              }
              .settings-liveviz:hover > .inner-settings-liveviz,.viz-types:hover > .inner-viz-types,.add-viz-data:hover > .add-viz-data-form{
                  visibility: visible;
                  opacity: 1;
              }
              .add-viz-data{
                  display:flex
              }
              .add-viz-data-form, .add-viz-data-form button{
                  margin-left: 10px;
              }
              .add-viz-data-form input {
                  width:50px;
              }
              .viz-types{
                  display: inline-flex;
              }
              .inner-viz-types{
                  margin-left: 10px;
              }
          </style>
              <script>
              var received_data;
              var current_graph_type = 1;
              window.onload = function() {
                  var parentDiv = document.getElementById("liveviz");
                  var c_width = parentDiv.clientWidth;
                  var c_height = parentDiv.clientHeight;
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
                              
                  // append the svg object to the body of the page
                  // append a 'group' element to 'svg'
                  // moves the 'group' element to the top left margin
                  var svg = d3.select("#liveviz").append("svg")
                      .attr("id","liveviz_svg")
                      .attr("width", width + margin.left + margin.right)
                      .attr("height", height + margin.top + margin.bottom)
                      .append("g")
                      .attr("transform", 
                              "translate(" + margin.left + "," + margin.top + ")");
                      var data = ${JSON.stringify(data)};
                      // format the data
                      data.forEach(function(d) {
                      d.y = +d.y;
                      });
                      received_data = data;
                      //received_data.sort((a, b) => d3.ascending(parseInt(a.x), parseInt(b.x)))
                      // Scale the range of the data in the domains
                      x.domain(data.map(function(d) { return d.x; }));
                      y.domain([0, d3.max(data, function(d) { return d.y; })]);
                  
                      // append the rectangles for the bar chart
                      svg.selectAll(".bar")
                          .data(data)
                      .enter().append("rect")
                          .attr("class", "bar")
                          .attr("x", function(d) { return x(d.x); })
                          .attr("width", x.bandwidth())
                          .attr("y", function(d) { return y(d.y); })
                          .attr("height", function(d) { return height - y(d.y); });
                  
                      // add the x Axis
                      svg.append("g")
                          .attr("transform", "translate(0," + height + ")")
                          .call(d3.axisBottom(x));
                  
                      // add the y Axis
                      svg.append("g")
                          .call(d3.axisLeft(y));
  
                      var f = document.createElement("form");
  
                      var keys = Object.keys(received_data[0]);
                      f.setAttribute('id',"addToGraph");
                      keys.forEach(element => {
                          var i = document.createElement("input"); //input element, text
                          i.setAttribute('type',"text");
                          i.setAttribute('name', element);    
                          i.setAttribute('placeholder', element);    
                          f.appendChild(i);              
                      })
      
                      var s = document.createElement("button"); //input element, Submit button
                      s.setAttribute('type',"button");
                      s.setAttribute('onclick',"addEntry()");
                      s.innerHTML = "ADD";
      
                      f.appendChild(s);
      
                      //and some more input elements here
                      //and dont forget to add a submit button
      
                      document.getElementsByClassName('add-viz-data-form')[0].appendChild(f);
              };
  
              function addEntry(){
                  var form = document.getElementById("addToGraph");
                  var keys = Object.keys(received_data[0]);
                  var allGood = true;
                  var new_entry = {};
                  keys.forEach(element => {
                      if(form.elements[element].value){
                          new_entry[element] = form.elements[element].value; 
                      }else{
                          alert("MISSING VALUE " + element);
                          allGood = false;
                      }
                  })
                  if(allGood && !received_data.find(element => element == new_entry)){
                      received_data.push(new_entry);
                      changeChartType(current_graph_type);
                  }
              }
          
              function changeChartType(type){
                  var parentDiv = document.getElementById("liveviz");
                  var c_width = parentDiv.clientWidth;
                  var c_height = parentDiv.clientHeight;
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
                      .attr("height", height + margin.top + margin.bottom)
                      .append("g")
                      .attr("transform", 
                              "translate(" + margin.left + "," + margin.top + ")");
                  
                  x.domain(received_data.map(function(d) { return d.x; }));
                  y.domain([0, d3.max(received_data, function(d) { return d.y; })]);
              
                  switch(type){
                      case 1:
                          current_graph_type = 1;
                          svg.selectAll(".bar")
                              .data(received_data)
                              .enter().append("rect")
                              .attr("class", "bar")
                              .style('cursor', 'pointer')
                              .attr("x", function(d) { return x(d.x); })
                              .attr("width", x.bandwidth())
                              .attr("y", function(d) { return y(d.y); })
                              .attr("height", function(d) { return height - y(d.y); })
                              .on("click", function(e) {
                                  received_data.splice(received_data.findIndex(v => v.x === e.x), 1);
                                  changeChartType(1);
                              });   
                          break;        
                      case 2:
                          current_graph_type = 2;
                          svg.selectAll(".dot")
                              .data(received_data)
                              .enter().append("circle")
                              .style("fill", "#000000")
                              .style('cursor', 'pointer')
                              .attr("cx", function(d) { return x(d.x) + x.bandwidth()/2; })
                              .attr("cy", function(d) { return y(d.y); })
                              .attr("r", 5)
                              .on("click", function(e) {
                                  received_data.splice(received_data.findIndex(v => v.x === e.x), 1);
                                  changeChartType(2);
                              });
                          break;
                      case 3:
                          current_graph_type = 3;
                          svg.append("path")
                              .datum(received_data)
                              .attr("fill", "none")
                              .attr("stroke", "#000000")
                              .attr("stroke-width", 3)
                              .attr("d", d3.line()
                                  .x(function(d) { return x(d.x) + x.bandwidth()/2; })
                                  .y(function(d) { return y(d.y) }))
                          svg.selectAll(".dot")
                              .data(received_data)
                              .enter().append("circle")
                              .style("fill", "#000000")
                              .style('cursor', 'pointer')
                              .attr("cx", function(d) { return x(d.x) + x.bandwidth()/2; })
                              .attr("cy", function(d) { return y(d.y); })
                              .attr("r", 5)
                              .on("click", function(e) {
                                  received_data.splice(received_data.findIndex(v => v.x === e.x), 1);
                                  changeChartType(3);
                              });
                          break;
                      case 4:
                          break;
                  }
                  // append the rectangles for the bar chart
              
                  // add the x Axis
                  svg.append("g")
                      .attr("transform", "translate(0," + height + ")")
                      .call(d3.axisBottom(x));
              
                  // add the y Axis
                  svg.append("g")
                      .call(d3.axisLeft(y));
              }
          </script>
          `;
        return result;
    }
}
module.exports = LiveViz;