const express = require("express");
const app = express();
const fs = require("fs");
const path = require("path");
const xlsxToJSON = require("xlsx-to-json");
const JSZip = require('jszip');
const Docxtemplater = require("docxtemplater");
const bodyParser = require("body-parser");
const mkdirp = require("mkdirp");
const handlebars  = require('express-handlebars');

app.engine('.hbs', handlebars({
  defaultLayout: 'main',
  extname: '.hbs'
}));
app.use(function(req, res, next) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "POST, PUT, GET, OPTIONS, DELETE");
  res.header("Access-Control-Max-Age", "3600");
  res.header(
    "Access-Control-Allow-Headers",
    "Content-Type, Access-Control-Allow-Headers, Authorization, X-Request"
  );
  next();
});
app.use(express.static(__dirname + '/styles'));
app.use(bodyParser.urlencoded({ extended: true }));

app.set('view engine', '.hbs');
app.set('views', path.join(__dirname + '/views'));

function readJSONFile(filename, callback) {
    fs.readFile(filename, function (err, data) {
      if(err) {
        callback(err);
        return;
      }
      try {
        callback(null, JSON.parse(data));
      } catch(exception) {
        callback(exception);
      }
    });
}

app.get("/", function(req, res) {
    res.render('index');
})

app.post("/",  function(req, res) {
  if(req.body.excelFile) {
    try {
       xlsxToJSON({
           input: req.body.excelFile,
           output: "input.json",
           // lowerCaseHeaders: true
         },
         function(err, result) {
           if (err) {
             res.end();
           } else {
             res.redirect("/generateWordFile")
           }
         }
       );
        } catch(err) {
          res.redirect("/error")
        }
  }
});

app.get("/generateWordFile", function(req, res) {
  readJSONFile('input.json', function(err, json) {
    if(err) throw err;
    res.render('generateFile', {input: json});
  })
})

app.post("/generateWordFile", function(req, res) {

    readJSONFile('input.json', function(err, json) {
        if(err) {
            throw err;
        }

        let folder = `./${req.body.folderName}`;
        mkdirp(folder, function(err) { 
            if(err) throw err;
        });

        try {
        for (let obj of json) {
           let content = fs.readFileSync(req.body.templateFile,
               "binary");

            let zip = new JSZip(content);
           
             let doc = new Docxtemplater();
             doc.loadZip(zip);
            
             //set the templateVariables
             doc.setData(obj);
            
             try {
               // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
               doc.render();
             } catch (error) {
               var e = {
                 message: error.message,
                 name: error.name,
                 stack: error.stack,
                 properties: error.properties
               };
               console.log(JSON.stringify({ error: e }));
               // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
               throw error;
             }
            
             var buf = doc.getZip().generate({ type: "nodebuffer" });
            
             // buf is a nodejs buffer, you can either write it to a file or do anything else with it.

             fs.writeFileSync(`./${req.body.folderName}/${Object.values(obj)[1]}.docx`, buf);
        }
        res.redirect("/generatedSuccessfully");
      }catch(err) {
        res.redirect("/error");
      }
    });
});

app.get("/generatedSuccessfully", function(req, res) {
  res.render('generatedSuccessfully.hbs');
})

app.get("/error", function(req, res) {
  res.render('error.hbs');
})

app.get("*", function(req, res) {
  res.status(400).render('error.hbs');
})

app.listen(3000);
