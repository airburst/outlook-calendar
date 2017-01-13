// Express for APIs
var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var o = require('./outlook');

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
var port = process.env.PORT || 8001;

var router = express.Router();

// Allow CORS (Testing only)
router.use(function (req, res, next) {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    next();
});


router.get('/api/calendar', o.calendarApi);

app.use('/', router);

// Start the server
app.listen(port);
console.log('Outlook API listening on port ' + port);