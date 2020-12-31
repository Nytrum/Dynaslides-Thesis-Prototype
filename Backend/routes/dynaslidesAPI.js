var express = require("express");
var router = express.Router();
const Dynaslides = require("../dynaslides");
const multer = require('multer');
const upload = multer();
router.post("/", upload.any() ,function(req, res) {
    var pptxfile = req.files[0];
    var args = {data_files: [], ai_on: false};
    for (let i = 1; i < req.files.length; i++) {
        args.data_files.push(req.files[i]); 
    }
    args.ai_on = (req.body["ai_on"] === 'true');
    const ds = new Dynaslides(pptxfile,args);
    ds.startDynaslides().then(function (slides){
        if(!ds.args.ai_on){
            var result_zip = ds.getCompressedFile(slides);
        }else{
            var result_zip = ds.getCompressedFileWithAudienceInteraction(slides);
        }
        result_zip
          .generateNodeStream({ type: "nodebuffer", streamFiles: true })
          .pipe(res)
          .on("end", function () {
            res.end();
        });
    });
});

module.exports = router;