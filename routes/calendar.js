var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* GET /calendar */
router.get('/', async function(req, res, next) {
  let parms = { title: 'Calendar', active: { calendar: true } };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // Set start of the calendar view to today at midnight
    const start = new Date(new Date().setHours(0,0,0));
    // Set end of the calendar view to 30 days from start
    const end = new Date(new Date(start).setDate(start.getDate() + 30));

    var Request = require("request");
    try {
      Request.post({
          "headers": { "content-type": "application/json", "Authorization":"Bearer "+accessToken},
          "url": "https://graph.microsoft.com/v1.0/me/findMeetingTimes",
          "body": JSON.stringify({
          "attendees": [
            {
              "emailAddress": {
                "address": "huynh@fabbier.onmicrosoft.com",
                "name": "Alex Darrow"
              },
              "type": "Required"
            }
          ],
          "timeConstraint": {
            "timeslots": [
              {
                "start": {
                  "dateTime": "2019-04-18T08:06:15.339Z",
                  "timeZone": "Pacific Standard Time"
                },
                "end": {
                  "dateTime": "2019-04-25T08:06:15.339Z",
                  "timeZone": "Pacific Standard Time"
                }
              }
            ]
          },
          "locationConstraint": {
            "isRequired": "false",
            "suggestLocation": "true",
            "locations": [
              {
                "displayName": "Conf Room 32/1368",
                "locationEmailAddress": "conf32room1368@imgeek.onmicrosoft.com"
              }
            ]
          },
          "meetingDuration": "PT1H"
        })
      }, (error, response, body) => {
          if(error) {
              return console.dir(error);
          }
          console.dir(response);
      });

    } catch(e) {
      console.log(e);
    }
    
  } else {
    // Redirect to home
    res.redirect('/');
  }
});

module.exports = router;