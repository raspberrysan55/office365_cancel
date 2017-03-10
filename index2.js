var server = require("./server2");
var router = require("./router");
var requestHandlers = require("./requestHandlers2");

var handle = {};
handle["/"] = requestHandlers.start;
handle["/start"] = requestHandlers.start;
handle["/upload"] = requestHandlers.upload;
handle["/post"] = requestHandlers.post;

server.start(router.route, handle);