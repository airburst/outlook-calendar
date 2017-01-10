var outlook = require('node-outlook');
var moment = require('moment');
var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/calendar'] = calendar;
handle['/new'] = addEvent;

server.start(router.route, handle);

// Home Route
function home(response, request) {
    console.log('Request handler \'home\' was called.');
    response.writeHead(200, { 'Content-Type': 'text/html' });
    response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
    response.end();
}

// Authorise Redirect Route
var url = require('url');
function authorize(response, request) {
    var url_parts = url.parse(request.url, true);
    var code = url_parts.query.code;
    authHelper.getTokenFromCode(code, tokenReceived, response);
}

// Save token as cookies
function tokenReceived(response, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
        response.writeHead(200, { 'Content-Type': 'text/html' });
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
    } else {
        getUserEmail(token.token.access_token, function (error, email) {
            if (error) {
                console.log('getUserEmail returned an error: ' + error);
                response.write('<p>ERROR: ' + error + '</p>');
                response.end();
            } else if (email) {
                var cookies = ['outlook-token=' + token.token.access_token + ';Max-Age=4000',
                'outlook-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                'outlook-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                'outlook-email=' + email + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                response.writeHead(302, { 'Location': 'http://localhost:8000/calendar' });
                response.end();
            }
        });
    }
}

function getUserEmail(token, callback) {
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
    var queryParams = {
        '$select': 'DisplayName, EmailAddress',
    };

    outlook.base.getUser({ token: token, odataParams: queryParams }, function (error, user) {
        if (error) {
            callback(error, null);
        } else {
            callback(null, user.EmailAddress);
        }
    });
}

function getValueFromCookie(valueName, cookie) {
    if (cookie.indexOf(valueName) !== -1) {
        var start = cookie.indexOf(valueName) + valueName.length + 1;
        var end = cookie.indexOf(';', start);
        end = end === -1 ? cookie.length : end;
        return cookie.substring(start, end);
    }
}

function getAccessToken(request, response, callback) {
    var expiration = new Date(parseFloat(getValueFromCookie('outlook-token-expires', request.headers.cookie)));

    if (expiration <= new Date()) {
        console.log('TOKEN EXPIRED, REFRESHING');
        var refresh_token = getValueFromCookie('outlook-refresh-token', request.headers.cookie);
        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {
            if (error) {
                callback(error, null);
            } else if (newToken) {
                var cookies = ['outlook-token=' + newToken.token.access_token + ';Max-Age=4000',
                'outlook-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                'outlook-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                callback(null, newToken.token.access_token);
            }
        });
    } else {
        // Return cached token
        var access_token = getValueFromCookie('outlook-token', request.headers.cookie);
        callback(null, access_token);
    }
}

// Calendar Route
function calendar(response, request) {
    getAccessToken(request, response, function (error, token) {
        var email = getValueFromCookie('outlook-email', request.headers.cookie);
        if (token) {
            response.writeHead(200, { 'Content-Type': 'text/html' });
            response.write('<div><h1>Your calendar</h1></div>');

            // TODO: accept start date from route
            var startDateStringUtc = moment().toISOString();
            var endDateStringUtc = moment().add(1, 'months').toISOString();
            var queryParams = {
                '$select': 'Subject,Start,End,IsAllDay',
                'startDateTime': startDateStringUtc,
                'endDateTime': endDateStringUtc,
                '$top': 100
            };

            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
            outlook.base.setAnchorMailbox(email);
            outlook.base.setPreferredTimeZone('Europe/London');

            outlook.calendar.getCalendarView({ token: token, odataParams: queryParams },
                function (error, result) {
                    if (error) {
                        console.log('getEvents returned an error: ' + error);
                        response.write('<p>ERROR: ' + error + '</p>');
                        response.end();
                    } else if (result) {
                        response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>All Day</th></tr>');
                        result.value.forEach(function (event) {
                            response.write(
                                '<tr><td>' + event.Subject +
                                '</td><td>' + moment(event.Start.DateTime).format('DD/MM/YYYY HH:mm') +
                                '</td><td>' + moment(event.End.DateTime).format('DD/MM/YYYY HH:mm') +
                                '</td><td>' + ((event.IsAllDay) ? 'Yes' : '') + '</td></tr>'
                            );
                        });
                        response.write('</table>');
                        response.end();
                    }
                });
        } else {
            response.writeHead(200, { 'Content-Type': 'text/html' });
            response.write('<p> No token found in cookie!</p>');
            response.end();
        }
    });
}

// Test method to create an event
function addEvent(response, request) {
    getAccessToken(request, response, function (error, token) {
        var email = getValueFromCookie('outlook-email', request.headers.cookie);
        if (token) {
            response.writeHead(200, { 'Content-Type': 'text/html' });
            response.write('<div><h1>Your calendar</h1></div>');

            var timezone = 'Europe/London';
            var event = {
                'Subject': 'Test Event using Integration',
                'Body': {
                    'ContentType': 'HTML',
                    'Content': 'A test message inside the event'
                },
                'Start': {
                    'DateTime': '2017-01-21T00:00:00',
                    'TimeZone': timezone
                },
                'End': {
                    'DateTime': '2017-01-23T00:00:00',
                    'TimeZone': timezone
                },
                'IsAllDay': false
            };

            // var queryParams = { 'event': event };

            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
            outlook.base.setAnchorMailbox(email);
            outlook.base.setPreferredTimeZone(timezone);

            outlook.calendar.createEvent({ token: token, event: event },
                function (error, result) {
                    if (error) {
                        console.log('createEvent returned an error: ' + error);
                        response.write('<p>ERROR: ' + error + '</p>');
                        response.end();
                    } else if (result) {
                        response.write('Event added');
                        response.end();
                    }
                });
        } else {
            response.writeHead(200, { 'Content-Type': 'text/html' });
            response.write('<p> No token found in cookie!</p>');
            response.end();
        }
    });
}