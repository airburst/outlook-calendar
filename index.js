var server = require('./server');
var o = require('./outlook');

function route(handle, pathname, response, request) {
    console.log('About to route a request for ' + pathname);
    if (typeof handle[pathname] === 'function') {
        return handle[pathname](response, request);
    } else {
        console.log('No request handler found for ' + pathname);
        response.writeHead(404, { 'Content-Type': 'text/plain' });
        response.write('404 Not Found');
        response.end();
    }
}

var handle = {};
handle['/'] = o.home;
handle['/authorize'] = o.authorize;
handle['/calendar'] = o.calendar;
// handle['/new'] = o.addEvent;

server.start(route, handle);