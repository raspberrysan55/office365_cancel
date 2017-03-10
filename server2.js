var http = require('http');
var https = require('https');
var url = require('url');
var qs = require('querystring');
var id = '';

function start(route, handle) {
  function onRequest(req, res) {
  // This time we don't store access token,
  // but please store access token and reuse in production code...
  var query = url.parse(req.url, true).query;
  var pathname = url.parse(req.url).pathname;

  if('code' in query) {
    // Get access token
    getAccessToken(query.code, function(jsonAuth) {
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
          res.write(msg.subject + '<br />' + msg.id);
          id = msg.id;
        }
       res.end();
      });
    });
 //   res.end();
  }
  else {
    // Redirect to login
    res.writeHead(302, {
      'Location':
        'https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&client_id=ea4efe0b-144e-4b4d-8b2c-6ae5985753c0&resource='
          + encodeURIComponent('https://graph.microsoft.com/')
          + '&redirect_uri='
          + encodeURIComponent('http://office365cancel.azurewebsites.net/')
    });
   res.end();    
  }
  
   var postData = '';
    console.log('Request for ' + pathname + 'received.');

    req.setEncoding("utf8");

    req.addListener("data", function(postDataChunk) {
      postData += postDataChunk;
      console.log("Received POST data chunk '"+
      postDataChunk + "'.");
    });

    req.addListener("end", function() {
      route(handle, pathname, res, postData);
    });

}
http.createServer(onRequest).listen(process.env.PORT);
  console.log("Server has started.");
}

exports.start = start;

function getAccessToken(code, callback) {
  var postdata = qs.stringify({
    'grant_type' : 'authorization_code',
    'code' : code,
    'client_id' : 'ea4efe0b-144e-4b4d-8b2c-6ae5985753c0',
    'client_secret' : '7eZ0ko8lXAJRKip6q4IXQUcQdH+krEXizkyrW7LQaRY=',
    'redirect_uri' : 'http://office365cancel.azurewebsites.net/'
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

function deleteEvent(access_token, datacallback, endcallback) {
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
