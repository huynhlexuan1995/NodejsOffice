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

    const event = {
        subject : "Let's go for lunch",
        body: {
          contentType: "HTML",
          content: "Does late morning work for you?"
        },
        start: {
            dateTime: "2017-04-15T12:00:00",
            timeZone: "Pacific Standard Time"
        },
        end: {
            dateTime: "2017-04-15T14:00:00",
            timeZone: "Pacific Standard Time"
        },
        location:{
            displayName:"Harry's Bar"
        },
        attendees: [
          {
            emailAddress: {
              address:"samanthab@contoso.onmicrosoft.com",
              name: "Samantha Booth"
            },
            type: "required"
          }
        ]
      }

    try {
      // Get the 10 newest messages from inbox
      const result = await client.api('/me/events').header({"content-type" : "application/json","Authorization":"Bearer"+accessToken}).post({ calendar:event });
      console.log(result);
    } catch (err) {
        parms.events = 'Error retrieving messages';
        parms.error = { status: `${err.code}: ${err.events}` };
        parms.debug = JSON.stringify(err.body, null, 2);
        res.render('error', parms);
    }
  } else {
    // Redirect to home
    res.redirect('/');
  }
});

module.exports = router;