class LiveCode {
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
        result += `<span class="plugin-name">Live code</span>`;
        result += `<div class="card"><pre style="overflow:auto;"><code contenteditable="true" class="javascript hljs" style="text-align:left;font-family:monospace !important;">`;
        this.text = this.text.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
        var parCount = 1;
        for (let i = 0; i < this.text.length - 1; i++) {
          if (this.text[i] == "{") {
            parCount += 1;
          }
          if (this.text[i] == "}") {
            parCount -= 1;
          }
          if (this.text[i] == "}" && parCount == 0) {
            break;
          }
          result += this.text[i];
        }
        result += `</code></pre></div><button class="run-js-button" onclick="runJavascriptCode(this);">Run code</button><div class="container-fluid result-container"><table><tbody class="code-result"></tbody></table></div>
        <style>
          .run-js-button{
            margin: 15px;
            padding: 0.2em 1.3em;
            cursor: pointer;
            color: #${this.colors[1] != undefined ? this.colors[1]: "fff"};
            border: none;
            background: #${this.colors[0] != undefined ? this.colors[0]: "000"};
          }
        </style>`;
        return result;
    }
}
module.exports = LiveCode;