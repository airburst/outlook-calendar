var outlook = require('node-outlook');
var moment = require('moment');
var authHelper = require('./authHelper');

// Home Route
function home(res, req) {
    console.log('req handler \'home\' was called.');
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
    res.end();
}

// Authorise Redirect Route
var url = require('url');
function authorize(res, req) {
    var url_parts = url.parse(req.url, true);
    var code = url_parts.query.code;
    authHelper.getTokenFromCode(code, tokenReceived, res);
}

// Save token as cookies
function tokenReceived(res, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.write('<p>ERROR: ' + error + '</p>');
        res.end();
    } else {
        getUserEmail(token.token.access_token, function (error, email) {
            if (error) {
                console.log('getUserEmail returned an error: ' + error);
                res.write('<p>ERROR: ' + error + '</p>');
                res.end();
            } else if (email) {
                var cookies = ['outlook-token=' + token.token.access_token + ';Max-Age=4000',
                'outlook-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                'outlook-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                'outlook-email=' + email + ';Max-Age=4000'];
                res.setHeader('Set-Cookie', cookies);
                res.writeHead(302, { 'Location': 'http://localhost:8000/calendar' });
                res.end();
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

function getAccessToken(req, res, callback) {
    var expiration = new Date(parseFloat(getValueFromCookie('outlook-token-expires', req.headers.cookie)));

    if (expiration <= new Date()) {
        console.log('TOKEN EXPIRED, REFRESHING');
        var refresh_token = getValueFromCookie('outlook-refresh-token', req.headers.cookie);
        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {
            if (error) {
                callback(error, null);
            } else if (newToken) {
                var cookies = ['outlook-token=' + newToken.token.access_token + ';Max-Age=4000',
                'outlook-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                'outlook-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
                res.setHeader('Set-Cookie', cookies);
                callback(null, newToken.token.access_token);
            }
        });
    } else {
        // Return cached token
        var access_token = getValueFromCookie('outlook-token', req.headers.cookie);
        callback(null, access_token);
    }
}

// Calendar Route
function calendar(res, req) {
    getAccessToken(req, res, function (error, token) {
        var email = getValueFromCookie('outlook-email', req.headers.cookie);
        if (token) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.write('<div><h1>Your calendar</h1></div>');

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
                        res.write('<p>ERROR: ' + error + '</p>');
                        res.end();
                    } else if (result) {
                        res.write('<table><tr><th>Subject</th><th>Start</th><th>End</th><th>All Day</th></tr>');
                        result.value.forEach(function (event) {
                            res.write(
                                '<tr><td>' + event.Subject +
                                '</td><td>' + moment(event.Start.DateTime).format('DD/MM/YYYY HH:mm') +
                                '</td><td>' + moment(event.End.DateTime).format('DD/MM/YYYY HH:mm') +
                                '</td><td>' + ((event.IsAllDay) ? 'Yes' : '') + '</td></tr>'
                            );
                        });
                        res.write('</table>');
                        res.end();
                    }
                });
        } else {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.write('<p> No token found in cookie!</p>');
            res.end();
        }
    });
}

// Test method to create an event
function addEvent(res, req) {
    getAccessToken(req, res, function (error, token) {
        var email = getValueFromCookie('outlook-email', req.headers.cookie);
        if (token) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.write('<div><h1>Your calendar</h1></div>');

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

            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
            outlook.base.setAnchorMailbox(email);
            outlook.base.setPreferredTimeZone(timezone);

            outlook.calendar.createEvent({ token: token, event: event },
                function (error, result) {
                    if (error) {
                        console.log('createEvent returned an error: ' + error);
                        res.write('<p>ERROR: ' + error + '</p>');
                        res.end();
                    } else if (result) {
                        res.write('Event added');
                        res.end();
                    }
                });
        } else {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.write('<p> No token found in cookie!</p>');
            res.end();
        }
    });
}

//======= APIs =============================================//
function handleError(res, reason, message, code) {
    console.log('API ERROR: ' + reason);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.write('{ "error": ' + message + '}');
    res.end();
}

// Calendar API Route
function calendarApi(res, req) {
    console.log( req.headers.cookie);   //
    getAccessToken(req, res, function (error, token) {
        var email = getValueFromCookie('outlook-email', req.headers.cookie);
        if (token) {
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
                        handleError(res, 'getCalendarView error', error, 400);
                    } else if (result) {
                        var events = [];
                        result.value.forEach(function (event) {
                            events.push({
                                id: event.Id,
                                title: event.Subject,
                                start: event.Start.DateTime,
                                end: event.End.DateTime,
                                allDay: event.IsAllDay
                            });
                        });
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.write('{"data":' + JSON.stringify(events) + '}');
                        res.end();
                    }
                });
        } else {
            handleError(res, 'Auth error', 'No token found in cookie', 400);
        }
    });
}

module.exports = {
    home: home,
    authorize: authorize,
    tokenReceived: tokenReceived,
    getUserEmail: getUserEmail,
    getValueFromCookie: getValueFromCookie,
    getAccessToken: getAccessToken,
    calendar: calendar,
    addEvent: addEvent,
    calendarApi: calendarApi
};