var outlook = require('node-outlook');
var moment = require('moment');
var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
// handle['/mail'] = mail;
handle['/calendar'] = calendar;

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
                var cookies = ['node-tutorial-token=' + token.token.access_token + ';Max-Age=4000',
                'node-tutorial-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                'node-tutorial-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                'node-tutorial-email=' + email + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                response.writeHead(302, { 'Location': 'http://localhost:8000/calendar' });
                response.end();
            }
        });
    }
}

function getUserEmail(token, callback) {
    // Set the API endpoint to use the v2.0 endpoint
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

    // Set up oData parameters
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
    var expiration = new Date(parseFloat(getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)));

    if (expiration <= new Date()) {
        // refresh token
        console.log('TOKEN EXPIRED, REFRESHING');
        var refresh_token = getValueFromCookie('node-tutorial-refresh-token', request.headers.cookie);
        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {
            if (error) {
                callback(error, null);
            } else if (newToken) {
                var cookies = ['node-tutorial-token=' + newToken.token.access_token + ';Max-Age=4000',
                'node-tutorial-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                'node-tutorial-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
                response.setHeader('Set-Cookie', cookies);
                callback(null, newToken.token.access_token);
            }
        });
    } else {
        // Return cached token
        var access_token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
        callback(null, access_token);
    }
}

// // Mail Route
// function mail(response, request) {
//     getAccessToken(request, response, function (error, token) {
//         // console.log('Token found in cookie: ', token);
//         var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
//         if (token) {
//             response.writeHead(200, { 'Content-Type': 'text/html' });
//             response.write('<div><h1>Your inbox</h1></div>');

//             var queryParams = {
//                 '$select': 'Subject,ReceivedDateTime,From,IsRead',
//                 '$orderby': 'ReceivedDateTime desc',
//                 '$top': 10
//             };

//             // Set the API endpoint to use the v2.0 endpoint
//             outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
//             // Set the anchor mailbox to the user's SMTP address
//             outlook.base.setAnchorMailbox(email);

//             outlook.mail.getMessages({ token: token, odataParams: queryParams },
//                 function (error, result) {
//                     if (error) {
//                         console.log('getMessages returned an error: ' + error);
//                         response.write('<p>ERROR: ' + error + '</p>');
//                         response.end();
//                     } else if (result) {
//                         // console.log('getMessages returned ' + result.value.length + ' messages.');
//                         response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
//                         result.value.forEach(function (message) {
//                             // console.log('  Subject: ' + message.Subject);
//                             var from = message.From ? message.From.EmailAddress.Name : 'NONE';
//                             response.write('<tr><td>' + from +
//                                 '</td><td>' + (message.IsRead ? '' : '<b>') + message.Subject + (message.IsRead ? '' : '</b>') +
//                                 '</td><td>' + message.ReceivedDateTime.toString() + '</td></tr>');
//                         });
//                         response.write('</table>');
//                         response.end();
//                     }
//                 });
//         } else {
//             response.writeHead(200, { 'Content-Type': 'text/html' });
//             response.write('<p> No token found in cookie!</p>');
//             response.end();
//         }
//     });
// }

// Calendar Route
function calendar(response, request) {
    getAccessToken(request, response, function (error, token) {
        var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
        if (token) {
            response.writeHead(200, { 'Content-Type': 'text/html' });
            response.write('<div><h1>Your calendar</h1></div>');

            var startDateStringUtc = moment().toISOString();
            var endDateStringUtc = moment().add(1, 'months').toISOString();
            var queryParams = {
                '$select': 'Subject,Start,End',
                '$startDateTime': startDateStringUtc,
                '$endDateTime': endDateStringUtc,
                '$orderby': 'Start/DateTime'
            };

            // Set the API endpoint to use the v2.0 endpoint
            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
            outlook.base.setAnchorMailbox(email);
            outlook.base.setPreferredTimeZone('Europe/London');

            outlook.calendar.getEvents({ token: token, odataParams: queryParams },
                function (error, result) {
                    if (error) {
                        console.log('getEvents returned an error: ' + error);
                        response.write('<p>ERROR: ' + error + '</p>');
                        response.end();
                    } else if (result) {
                        // console.log('getEvents returned ' + result.value.length + ' events.');
                        response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th></tr>');
                        result.value.forEach(function (event) {
                            response.write('<tr><td>' + event.Subject +
                                '</td><td>' + moment(event.Start.DateTime).format('DD/MM/YYYY HH:mm') +
                                '</td><td>' + moment(event.End.DateTime).format('DD/MM/YYYY HH:mm') + '</td></tr>');
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