var https = require('https');
var Express = require('express');
var multer = require('multer');
var cmd =  require('node-cmd');
var path = require('path');
var fs = require('fs');
var bodyParser = require('body-parser');
var app = Express();
var options = {
key: fs.readFileSync("privatekey.pem"),
cert: fs.readFileSync("certificate.pem")	
};

app.use(bodyParser.json());
app.use(Express.static('public'));

app.get('/', function (req, res){
  res.sendFile( __dirname  + '/public/powerpoint.html');
});

var Storage = multer.diskStorage({
  destination: function (req, file, callback) {
    callback(null, "./public");
    
  },
 
  filename: function (req, file, callback) {
     
    var first = file.originalname.split(".")[0];
    
   
    var last = file.originalname.split(".")[1];
     var file =  first  + Date.now();
    var fileName =   first  + Date.now()+'.'+last;
   
    callback(null,fileName);
    
    cmd.run('\public.\\controll.exe \\' +fileName); 
  },
 
 
});


var upload = multer({ storage: Storage });



app.post("/upload",upload.single('file'),function (req, res) {
 // var fileName = req.file.filename;
  var fileName = req.file.filename.split('.')[0]
 
 res.send(fileName);

})

app.get('/upload',function(req,res){
	 
	res.send('');
})



https.createServer(options,app, function(req,res) {
	res.writeHead(200);
	res.end();	
}).listen(2000);
