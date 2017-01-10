require('dotenv').config();

var appId = process.env.APP_ID;
var secret = process.env.APP_SECRET;
var redirectUri = process.env.REDIRECT_URI;
var credentials = {
    client: {
        id: appId,
        secret: secret,
    },
    auth: {
        tokenHost: 'https://login.microsoftonline.com',
        authorizePath: 'common/oauth2/v2.0/authorize',
        tokenPath: 'common/oauth2/v2.0/token'
    }
};

var oauth2 = require('simple-oauth2').create(credentials);

// The scopes the app requires
var scopes = [
    'openid',
    'offline_access',
    'https://outlook.office.com/mail.read',
    'https://outlook.office.com/calendars.readwrite'
];

function getAuthUrl() {
    var returnVal = oauth2.authorizationCode.authorizeURL({
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    });
    return returnVal;
}

exports.getAuthUrl = getAuthUrl;

function getTokenFromCode(auth_code, callback, response) {
    var token;
    oauth2.authorizationCode.getToken({
        code: auth_code,
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    }, function (error, result) {
        if (error) {
            console.log('Access token error: ', error.message);
            callback(response, error, null);
        } else {
            token = oauth2.accessToken.create(result);
            console.log('Token created: ', token.token);
            callback(response, null, token);
        }
    });
}

exports.getTokenFromCode = getTokenFromCode;

function refreshAccessToken(refreshToken, callback) {
    var tokenObj = oauth2.accessToken.create({ refresh_token: refreshToken });
    tokenObj.refresh(callback);
}

exports.refreshAccessToken = refreshAccessToken;