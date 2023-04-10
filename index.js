
const express = require("express")
const app = express();
var bodyParser = require('body-parser');

var json2xls = require('json2xls');
const fs = require("fs");
const multer = require("multer");
const path = require('path');
const cors = require('cors');

const csv = require("csvtojson");
const excelToJson = require('convert-excel-to-json');

app.use(cors())
app.use(express.static('public'));

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());


app.set('view engine', 'ejs');

app.get("/", function (req, res) {
    res.render('upload');
});


var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/')
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + "__" + file.originalname)
    }
});

var upload = multer({ storage: storage }).single('filename');

let finalData;
let filename;
let templatedata;
app.post("/upload", upload, function (req, res) {
    templatedata = JSON.parse(fs.readFileSync(__dirname + '\\' + 'excelConfig.json'))

    // console.log(templatedata);
    filename = req.file.filename
    let file = filename.split('.')
    let fileextn = file[file.length - 1]
    // console.log(fileextn);

    let filepath = path.join(__dirname, '/uploads', filename)
    // console.log(filepath);

    /*to read csv file */

    if (fileextn === 'csv') {
        async function readcsvFile(filename) {
            csv()
                .fromFile(filename)
                .then(async (jsonObj) => {
                    // console.log(jsonObj);
                    finalData = jsonObj
                    res.render('result', { data: jsonObj })
                })
                .catch((err) => {
                    console.log(err.message);
                });
        }
        readcsvFile(filepath);
    }
    else if (fileextn === 'xlsx') {
        let jsonObj = excelToJson({
            sourceFile: filepath,
            columnToKey: {
                "*": "{{columnHeader}}"
            }
        })
        console.log(jsonObj,"json obj");
        console.log(jsonObj.Sheet1,"sheet1");
        finalData = jsonObj;
        res.render('result', { data: jsonObj.Sheet1, headers: templatedata })
    }

    else {
        res.render('error')
    }

});


app.get("/edit", upload, function (req, res) {
    let val = req.query.value
    // console.log(val);
    var templatedata = JSON.parse(fs.readFileSync(__dirname + '\\' + 'excelConfig.json'))

    console.log(finalData,"sheet1 array");
   if(val){
    console.log(templatedata);
    res.render('edit', { data: finalData.Sheet1[val],val:val, headers: templatedata })
   }
   else{
    res.render('new',{ headers:templatedata})
   }
});


app.get('/upload', (req, res) => {
    if(finalData){
        res.render('result', { data: finalData.Sheet1, headers: templatedata });
    }
})

app.get('/save', (req, res) => {
    var xls = json2xls(finalData.Sheet1);
    
fs.writeFileSync('data.xlsx', xls, 'binary');
})

app.use(json2xls.middleware);


app.post("/form",function(req,res){
   if(req.query.value){
    let val = req.query.value;
    console.log(req.body,"edit form ","val :",val);

        // 'Variable 1': '13',
        // 'Variable 2': 'CL',
        // 'Variable 3': 'po.entry',
        // 'Variable 4': 'showedi',
        // 'Variable 5': 'jobr',
        // 'Variable 6': '00108',
        // 'Variable 7': '',
        // 'Variable 8': '10',
        // 'Variable 9': '20',
        // 'Variable 10': '25',
        // 'Variable 11': '',
        // 'Variable 12': 'EDI',
        // 'Variable 13': '65450',
        // Expected_Result: 'EDI'
    finalData.Sheet1[val].TestCase_ID = req.body.TestCase_ID;
    finalData.Sheet1[val].Test_Case_Title= req.body.Test_Case_Title;
    finalData.Sheet1[val].Execution= req.body.Execution;
    finalData.Sheet1[val].Environment= req.body.Environment;
    finalData.Sheet1[val].Database = req.body.Database;
    finalData.Sheet1[val].Scenario_ID= req.body.Scenario_ID;
    finalData.Sheet1[val].Databasenumber= req.body.Databasenumber;
    finalData.Sheet1[val].Entity= req.body.Entity;
    finalData.Sheet1[val].Expected_Result= req.body.Expected_Result;
    
    res.render('result', { data: finalData.Sheet1, headers: templatedata});
   }
   else{
    finalData.Sheet1.push(req.body);
    console.log(finalData,"final data after new record");
    console.log(req.body,"body from form");
    // res.xls('data.xlsx', finalData.Sheet1);

    res.render('result', { data: finalData.Sheet1, headers: templatedata });
   }

  
  });



app.listen(3001, function () {
    console.log('Server running on port 3001');
});

