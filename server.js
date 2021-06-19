const express = require('express');
const app = express();
const xlsxtojson = require('xlsx-to-json');
const xlstojson = require("xls-to-json");
const crypto = require('crypto');
let fileExtension = require('file-extension');

app.use(express.json());
app.use(express.static("public"));
app.use(express.static("uploads"));
// set the view engine to ejs
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: false }));

//#region Multer upload
const multer = require('multer');
let storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './input/')
    },
    filename: function (req, file, cb) {
        crypto.pseudoRandomBytes(16, function (err, raw) {
            cb(null, raw.toString('hex') + Date.now() + '.' + fileExtension(file.mimetype));
            });
    }
});
let upload = multer({storage: storage}).single('file');
//#endregion


/** Method to handle the form submit */
app.post('/sendFile', async function(req, res) {
    let excel2json;
    upload(req,res,function(err){
        if(err){
             res.json({error_code:401,err_desc:err});
             return;
        }
        if(!req.file){
            res.json({error_code:404,err_desc:"File not found!"});
            return;
        }

        if(req.file.originalname.split('.')[req.file.originalname.split('.').length-1] === 'xlsx'){
            excel2json = xlsxtojson;
        } else {
            excel2json = xlstojson;
        }

       //  code to convert excel data to json  format
        excel2json({
            input: req.file.path,  
            output: "output/"+Date.now()+".json", // output json 
            lowerCaseHeaders:true
        }, async function(err, result) {
            if(err) {
              res.json(err);
            } else {
                var keys = [];
                if(result && result.length > 0){
                    var firstElem = result[0];
                    Object.keys(firstElem).forEach(element => {
                        keys.push(element);
                    });
                    res.render('templateViewer', { list: result, keys: keys });
                }
                else{
                    res.json(result);
                }
            }
        });
    })

});


app.get('/', (req, res) => {
    res.render('index.ejs')
});

app.listen(3000, () => {
    console.log('App running on port 3000');
})