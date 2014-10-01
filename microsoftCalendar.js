module.exports.OutlookCalendar = OutlookCalendar;
var util = require('util');
var request = require('request');
function OutlookCalendar(accessToken, refreshToken){
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
}

OutlookCalendar.prototype.updateToken = function(accessToken) {
    this.accessToken = accessToken;
};

OutlookCalendar.prototype.get = function (eventId, done){
    request('https://apis.live.net/v5.0/me/' + eventId + '?access_token=' + this.accessToken, function (error, response, body){
        if (error) done(error, null);
        else {
            done(null, JSON.parse(response.body));
        }
    });
};

OutlookCalendar.prototype.list = function (done){
    request('https://apis.live.net/v5.0/me/events?access_token=' + this.accessToken, function (error, response, body){
        if (error) done(error, null);
        else {
            done(null, JSON.parse(response.body));
        }
    });
};

OutlookCalendar.prototype.update = function (eventId, body, done){
    request({url: 'https://apis.live.net/v5.0/me/' + eventId + '?access_token=' + this.accessToken, method: "PATCH", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    });
};

OutlookCalendar.prototype.delete = function (eventId, done){
    request.delete('https://apis.live.net/v5.0/me/' + eventId + '?access_token=' + this.accessToken, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    });
};

OutlookCalendar.prototype.create = function (body, done){
    request({url: 'https://apis.live.net/v5.0/me/events?access_token=' + this.accessToken, method: "POST", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    })
};

OutlookCalendar.prototype.respond = function (eventID, response, body, done){
    request({url: 'https://apis.live.net/v5.0/me/' + eventID + '/' + response + '?access_token=' + this.accessToken, method: "POST", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    })
};

OutlookCalendar.prototype.refreshAccessToken = function (refreshToken, client_id, client_secret, resource, done){
    var _obj = {
        'grant_type': 'refresh_token',
        'refresh_token': refreshToken,
        'client_id': client_id,
        'client_secret': client_secret,
        'resource': resource
    };

    request({url: 'https://login.windows.net/common/oauth2/token', method: "POST", body: _obj}, function (error, response, body){
        done(error, response.body || "");
    });
};
