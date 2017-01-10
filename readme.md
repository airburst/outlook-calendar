#Outlook Calendar Integration

This is a small API to sign in to Office365 with a user context and then make the user's calendar available as a simple API.


##Installation

* Clone fron Github and then run the usual `yarn` to install dependencies
* Add the function below into `node_modules/node-outlook/calendar-apis.js`
* Add your own values into .env file in root for `APP_ID`, `APP_SECRET` and `REDIRECT_URL` 
* Run `npm start`

    
    // Monkey Patched into API because there is
    // no lib support for this REST method yet 
    getCalendarView: function (parameters, callback) {
        var userSpec = utilities.getUserSegment(parameters);

        var requestUrl = base.apiEndpoint() + userSpec + '/calendarview';

        var apiOptions = {
            url: requestUrl,
            token: parameters.token,
            user: parameters.user
        };

        if (parameters.odataParams !== undefined) {
            apiOptions['query'] = parameters.odataParams;
        }

        base.makeApiCall(apiOptions, function (error, response) {
            if (error) {
            if (typeof callback === 'function') {
                callback(error, response);
            }
            }
            else if (response.statusCode !== 200) {
            if (typeof callback === 'function') {
                callback('REST request returned ' + response.statusCode + '; body: ' + JSON.stringify(response.body), response);
            }
            }
            else {
            if (typeof callback === 'function') {
                callback(null, response.body);
            }
            }
        });
    }

