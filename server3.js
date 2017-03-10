var http = require('http');
var https = require('https');
var url = require('url');
var qs = require('querystring');
var moment = require("moment");
var id = "";
var time = "";

http.createServer(
  function onRequest (req, res) {
    // This time we don't store access token,
    // but please store access token and reuse in production code...
    var query = url.parse(req.url, true).query;
    console.log(query);
    
    var postData = "";
    req.setEncoding("utf8");
   if (req.method === 'POST' && req.url === '/') { 
    req.on("data", function(postDataChunk) {
      postData += postDataChunk;
      console.log("Received POST data chunk '"+
        postDataChunk + "'.");
        });
        
    req.on("end", function() {
       time = JSON.parse(postData).time;
       console.log("-----------"+ time + "---------");
       });
   }   
       // Get access token
       getAccessToken(function(jsonAuth) { 
       // Get messages from Office 365 (Exchange Online)
       res.writeHead(200,
         { 'Content-Type': 'text/html; charset=utf-8' });
         var msgbody = '';
         getEvent(JSON.parse(jsonAuth).access_token,
           function(jsonMsg) {
             msgbody += jsonMsg;
             }, function() {
               var msgobj = JSON.parse(msgbody).value;
               for(var i = 0; i < msgobj.length; i++) {
                 var msg = msgobj[i];
        
        
        
        res.write(msg.subject + '<br />' + msg.start.dateTime + '<br />' + msg.end.dateTime + '<br /><br />');
        id = msg.id;
      }
     res.end();
    });
    
 //センサーからPOSTがきたときだけDELETE   
if (req.method === 'POST' && req.url === '/') {
   deleteEvent(id, JSON.parse(jsonAuth).access_token,
function(jsonMsg) {
   }, function() {     
    });
}  
});
  
}).listen(process.env.PORT);
 
function getAccessToken(callback) {
  var postdata = qs.stringify({
      'grant_type' : 'password',
      'resource' : 'https://graph.microsoft.com/',
      'client_id' : 'ea4efe0b-144e-4b4d-8b2c-6ae5985753c0',
      'client_secret' : '7eZ0ko8lXAJRKip6q4IXQUcQdH+krEXizkyrW7LQaRY=',
      'username' : 'raspberrysan55@raspberrysan55.onmicrosoft.com',
      'password' : 'Raspberry3720'
  });
  var opt = {
    host : 'login.windows.net',
    port : 443,
    path : '/common/oauth2/token',
    method : 'POST',
    headers : {
      'Content-Type' : 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(postdata)
    }
  };
  var authreq = https.request(opt, function(authres) {
    authres.setEncoding('utf-8');
    authres.on('data', callback);
  });
  authreq.write(postdata);
  authreq.end();
}

function getEvent(access_token, datacallback, endcallback) {
  var opt = {
    host : 'graph.microsoft.com',
    port : 443,
    path : '/beta/me/events/',
    method : 'GET',
    headers : {
      'Authorization' : 'Bearer ' + access_token,
      'Content-Length': 0
    }
  };
  var o365req = https.request(opt, function(o365res) {
    o365res.setEncoding('utf-8');
    o365res.on('data', datacallback);
    o365res.on('end', endcallback);
  });
  o365req.end();
}

function deleteEvent(id,access_token, datacallback, endcallback) {
  var opt = {
    host : 'graph.microsoft.com',
    port : 443,
    path : '/beta/me/events/' + id,
    method : 'DELETE',
    headers : {
      'Authorization' : 'Bearer ' + access_token,
      'Content-Length': 0
    }
  };
  var o365req = https.request(opt, function(o365res) {
    o365res.setEncoding('utf-8');
    o365res.on('data', datacallback);
    o365res.on('end', endcallback);
  });
  o365req.end();
}