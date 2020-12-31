class LiveQuestion {
    constructor(command,text,colors) {
        this.command = command;
        this.text = text;
        this.colors = colors;
    }
    /**
    * This is where the html plugin funcionality is added.
    * @return
    *   The HTML in a text format
    */
    create(){
        var result = "";
        var data_question = this.command.match(/&quot;(.*?)&quot;/g);
      result += `<span class="plugin-name">Live question</span>`;
      result += `
        <div id="liveviz" style="height: 100%;">
        </div>
        <script>
        var received_data;
        var current_graph_type = 1;
        window.onload = function() {
          if(window.presentation_type == "client"){
            var parentDiv = document.getElementById("liveviz");
            var c_width = parentDiv.clientWidth;
            var c_height = parentDiv.clientHeight;
            var question_div = document.createElement("DIV");
            var buttons_div = document.createElement("DIV");
            buttons_div.setAttribute("id", "buttons_div")
            var voted_div = document.createElement("DIV");
            voted_div.setAttribute("id", "voted_div")
            var thanks_p = document.createElement("p");
            var update_p = document.createElement("p");
            var btn_yes = document.createElement("BUTTON");
            var btn_no = document.createElement("BUTTON");
            question_div.innerHTML = "${data_question[0].replace(/&quot;/g,'')}";
            thanks_p.innerHTML = "Thank you for voting!";
            update_p.innerHTML = "See the presenter's screen for voting update!";
            btn_yes.className = "voting-button";
            btn_no.className = "voting-button";
            btn_yes.innerHTML = "YES";
            btn_no.innerHTML = "NO";
            question_div.style = "padding-bottom: 1em; text-align:center; font-size: 3em; font-weight:700;";
            buttons_div.style = "width:100%;display:flex; justify-content:center; align-items:center;";
            voted_div.style = "display:none;";
            thanks_p.style = "padding: 5%;font-size: 2.5em;color: #${this.colors[0] != undefined ? this.colors[0]: "000"};";
            update_p.style = "font-size: 1.1em;color: #${this.colors[0] != undefined ? this.colors[0]: "000"}70;";
            btn_yes.style = "color: white; background-color: green;";
            btn_no.style = "color: white;background-color: red;";
            parentDiv.appendChild(question_div);            
            parentDiv.appendChild(buttons_div);            
            parentDiv.appendChild(voted_div);            
            voted_div.appendChild(thanks_p);            
            voted_div.appendChild(update_p);            
            buttons_div.appendChild(btn_yes);            
            buttons_div.appendChild(btn_no);   
            btn_yes.onclick = function(event) {
              var slideNumber = Reveal.getSlidePastCount();
              var data = {
                  type: 3,
                  slide:slideNumber,
                  message: 'YES'
              }
              sendQuestion(data);
              document.getElementById('voted_div').style.display = "block";
              document.getElementById('buttons_div').style.display = "none";
            }
            btn_no.onclick = function(event) {
              var slideNumber = Reveal.getSlidePastCount();
              var data = {
                  type: 3,
                  slide:slideNumber,
                  message: 'NO'
              }
              sendQuestion(data);
              document.getElementById('voted_div').style.display = "block";
              document.getElementById('buttons_div').style.display = "none";
            }       
          }else{
            var parentDiv = document.getElementById("liveviz");
            var c_width = parentDiv.clientWidth;
            var c_height = parentDiv.clientHeight* 0.8;
            var question_div = document.createElement("DIV");
            question_div.innerHTML = "${data_question[0].replace(/&quot;/g,'')}";
            question_div.style = "text-align: center;font-size: 2.5em;width: 100%;font-weight: 700;height:20%";
            parentDiv.insertBefore(question_div, parentDiv.firstChild);
            window.question_div_height = question_div.clientHeight;
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
                .attr("height", height + margin.top + margin.bottom + 30)
                .append("g")
                .attr("transform", 
                        "translate(" + margin.left + "," + (margin.top + 30) + ")");
                var data = [
                  {x: "YES", y: 0},
                  {x: "NO", y: 0}
                ];
                // format the data
                data.forEach(function(d) {
                d.y = +d.y;
                });
                window.received_data = data;
                //received_data.sort((a, b) => d3.ascending(parseInt(a.x), parseInt(b.x)))
                // Scale the range of the data in the domains

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
            
              };
        }
        function sendQuestion(data) {

          var messageData = {
              id:getCookieInner(window.presentation_name + "_id"),
              data: data,
              secret: multiplex.secret_c,
              socketId: multiplex.id_c
          };
        
        
          window.my_socket.emit( 'multiplex-statechanged', messageData );
        
        }

        function getCookieInner(cname) {
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

      </script>`
      return result;
    }
}
module.exports = LiveQuestion;