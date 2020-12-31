import React, { Component } from "react";

import axios from 'axios';

import FileDownload from 'js-file-download';

import "./App.css";

class App extends Component {

    constructor(props) {
        super(props);
        this.state = { 
            selectedPowerpointFile: null,
            selectedDataFiles: [],
            ai_on: false,
            dspresentation: null
        };
    }

    callDynaslidesAPI() {
        fetch("http://localhost:7000/getDynaslidesPresentation")
            .then(res => res.text())
            .then(res => console.log(res))
            .catch(err => err);
    }

    onChangeHandler = event => {
        this.resetDownloadAvailable();
        this.setState({
            selectedPowerpointFile: event.target.files[0]
        });
    }
    onChangeHandler2 = event => {
        if(event.target.checked){
            this.setState({
                ai_on: false
            });
        }else{
            this.setState({
                ai_on: true
            });
        }
    }

    onChangeHandler3 = event => {
        this.setState({
            selectedDataFiles: event.target.files
        });
    }

    onClickHandler = () => {
        const data = new FormData();
        var border_special = document.getElementsByClassName('border-special')[0];
        border_special.classList.add("visibility-visible");
        border_special.disabled = true;
        data.append('file', this.state.selectedPowerpointFile);
        for(var x = 0; x<this.state.selectedDataFiles.length; x++) {
            data.append('file', this.state.selectedDataFiles[x])
        }
        data.append('ai_on', this.state.ai_on);
        axios.post("http://localhost:7000/getDynaslidesPresentation", data, { 
            responseType: 'arraybuffer'
        }).then(res => { // then print response status
            this.setState({
                dspresentation: res.data
            });
            this.makeDownloadAvailable();
        })
    }

    onClickHandler2 = () => {
        FileDownload(this.state.dspresentation, this.state.selectedPowerpointFile.name.split('.').slice(0, -1).join('.') + '.zip');
    }

    onClickHandler3 = () => {
        var additional_settings = document.getElementById('additional-settings-div');
        if (additional_settings.classList.contains("visibility-visible")){
            additional_settings.classList.remove("visibility-visible");
            additional_settings.classList.remove("height-control");
        }else{
            additional_settings.classList.add("visibility-visible");
            additional_settings.classList.add("height-control");
        }
    }
    
    componentDidMount() {
        // appendScript("./js/jquery-3.4.1.min.js");
        // appendScript("./js/popper.min.js");
        // appendScript("./js/bootstrap.min.js");
        // // appendScript("./js/d3.min.js");
        // // appendScript("./js/nv.d3.min.js");
        // // appendScript("./js/EasePack.min.js");
        // // appendScript("./js/TweenLite.min.js");
        // // appendScript("./js/graph-animation.js");
    }

    makeDownloadAvailable(){
        var border_special = document.getElementsByClassName('border-special')[0];
        var border_special_inner = document.getElementsByClassName('border-special__inner')[0];
        var download_text = document.getElementsByClassName('download-text')[0];
        var d_spinner1 = document.getElementsByClassName('d-spinner__one')[0];
        var d_spinner2 = document.getElementsByClassName('d-spinner__two')[0];
        var d_spinner3 = document.getElementsByClassName('d-spinner__three')[0];
        var d_spinner4 = document.getElementsByClassName('d-spinner__four')[0];
        var download_button = document.getElementById('download-reveal-compressed-btn');
        border_special.disabled = false;
        border_special.classList.add("border-animation");
        border_special_inner.classList.add("next-background-border-special");
        border_special_inner.classList.add("border-animation__inner");
        download_text.classList.add("visibility-visible");
        d_spinner1.classList.add("stop-anim-count");
        d_spinner2.classList.add("stop-anim-count");
        d_spinner3.classList.add("stop-anim-count");
        d_spinner4.classList.add("stop-anim-count");
        download_button.classList.add("hoverable-boxshadow");        
    }

    resetDownloadAvailable(){
        var border_special = document.getElementsByClassName('border-special')[0];
        var border_special_inner = document.getElementsByClassName('border-special__inner')[0];
        var download_text = document.getElementsByClassName('download-text')[0];
        var d_spinner1 = document.getElementsByClassName('d-spinner__one')[0];
        var d_spinner2 = document.getElementsByClassName('d-spinner__two')[0];
        var d_spinner3 = document.getElementsByClassName('d-spinner__three')[0];
        var d_spinner4 = document.getElementsByClassName('d-spinner__four')[0];
        var download_button = document.getElementById('download-reveal-compressed-btn');
        border_special.classList.remove("border-animation");
        border_special.classList.remove("visibility-visible");
        border_special_inner.classList.remove("next-background-border-special");
        border_special_inner.classList.remove("border-animation__inner");
        download_text.classList.remove("visibility-visible");
        d_spinner1.classList.remove("stop-anim-count");
        d_spinner2.classList.remove("stop-anim-count");
        d_spinner3.classList.remove("stop-anim-count");
        d_spinner4.classList.remove("stop-anim-count");
        download_button.classList.remove("hoverable-boxshadow");        
    }

    render() {
        return (
            <div className="App">
                <div className="container">
                    <div className="dyna-header">
                        <h1 className="dyna-title">DynaSlides</h1>
                        <h4 className="dyna-slogan">Make your presentations more dynamic!</h4>
                    </div>
                    <div className="row">
                        <div className="col center-style">
                            <div className="interaction-options">
                                <div className="fileUpload custom-btn">
                                    <label>Choose a PPTX file</label>
                                    <input id="uploadBtn" type="file" className="upload wh100" onChange={this.onChangeHandler} accept="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
                                </div>
                                <button className="additional-settings" onClick={this.onClickHandler3}>Additional settings</button>
                            </div>
                        </div>
                        {/* <div className="col-md-4 or-span">OR</div>
                        <div className="col-md-4 center-style">
                            <button id="create-empty-presentation" className="custom-btn" ><i className="glyphicon glyphicon-export"></i>Create empty presentation</button>
                        </div> */}
                    </div>
                    <div id="additional-settings-div" className="row data-files-container visibility-hidden">
                        <div className="col-md-6 ">
                            <div className="fileUpload data-insertion-btn">
                                <label>Insert data files used for presentation</label>
                                <input id="uploadBtn2" onChange={this.onChangeHandler3} type="file"  className="upload wh100" accept=".csv, .json" multiple/>
                            </div>
                        </div>
                        <div className="col-md-6 ai-switch-button">
                            <p className="ai-switch-text">Audience Feedback</p>
                            <div className="checkbox">
                                <input type="checkbox" name="example" id="example" defaultChecked="true" onChange={this.onChangeHandler2}/>
                                <div className="checkbox-inner">
                                    <label htmlFor="example"></label><span></span>
                                </div>
                                <div className="checkbox__on">ON</div>
                                <div className="checkbox__off">OFF</div>
                            </div>
                        </div>
                    </div>
                    {this.state.selectedPowerpointFile != null &&
                        <div className="row">
                            <div className="col">
                                <button type="button" className="btn custom-btn upload-button" onClick={this.onClickHandler}>
                                    <svg className="upload-icon" x="0px" y="0px" width="15px" height="15px" viewBox="0 0 433.5 433.5">
                                        <g>
                                            <g id="file-upload">
                                                <polygon points="140.25,331.5 293.25,331.5 293.25,178.5 395.25,178.5 216.75,0 38.25,178.5 140.25,178.5"/>
                                                <rect x="38.25" y="382.5" width="357" height="51"/>
                                            </g>
                                        </g>
                                    </svg>
                                    Upload  
                                </button> 
                            </div>
                        </div>
                    }
                    <div id="download-interactions" className="row download-container">
                        <div className="col-md-12 center-style" >
                            <button id="download-reveal-compressed-btn" className="border-special" onClick={this.onClickHandler2}>
                                <div className="border-special__inner">
                                    <svg className="d-spinner-svg" width='150px' height='179px'>
                                        <path className='d-spinner d-spinner__four' d='M144.421372,121.923755 C143.963266,123.384111 143.471366,124.821563 142.945674,126.236112 C138.856723,137.238783 133.098899,146.60351 125.672029,154.330576 C118.245158,162.057643 109.358082,167.978838 99.0105324,172.094341 C89.2149248,175.990321 78.4098994,178.04219 66.5951642,178.25 L0,178.25 L144.421372,121.923755 L144.421372,121.923755 Z'/>
                                        <path className='d-spinner d-spinner__three' d='M149.033408,92.6053108 C148.756405,103.232477 147.219069,113.005232 144.421372,121.923755 L0,178.25 L139.531816,44.0158418 C140.776016,46.5834381 141.913968,49.2553065 142.945674,52.0314515 C146.681818,62.0847774 148.711047,73.2598899 149.033408,85.5570717 L149.033408,92.6053108 L149.033408,92.6053108 Z'/>
                                        <path className='d-spinner d-spinner__two' d='M80.3248924,1.15770478 C86.9155266,2.16812827 93.1440524,3.83996395 99.0105324,6.17322306 C109.358082,10.2887257 118.245158,16.2099212 125.672029,23.9369874 C131.224984,29.7143944 135.844889,36.4073068 139.531816,44.0158418 L0,178.25 L80.3248924,1.15770478 L80.3248924,1.15770478 Z'/>
                                        <path className='d-spinner d-spinner__one' d='M32.2942065,0 L64.5884131,0 C70.0451992,0 75.290683,0.385899921 80.3248924,1.15770478 L0,178.25 L0,0 L32.2942065,0 L32.2942065,0 Z'/>
                                    </svg>
                                    <span className="download-text">
                                        <svg className="download-icon" x="0px" y="0px" width="15px" height="15px" viewBox="0 0 433.5 433.5">
                                            <g>
                                                <g id="file-download">
                                                    <path d="M395.25,153h-102V0h-153v153h-102l178.5,178.5L395.25,153z M38.25,382.5v51h357v-51H38.25z"/>
                                                </g>
                                            </g>
                                        </svg>
                                        Download  
                                    </span>
                                </div>
                            </button>
                        </div>
                        {/* <div className="col-md-2 or-span center-style" >
                            OR
                        </div>
                        <div className="col-md-5 center-style">
                            <button id="download-reveal-single-file-btn" className="custom-btn" ><i className="glyphicon glyphicon-export"></i>Download as a single file</button>
                        </div> */}
                    </div>
                </div>
            </div>
        );
    }
}

export default App;