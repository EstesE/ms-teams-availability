var counter = 1;
async function msTeams() {
    var token = localStorage["adal.access.token.keyhttps://api.spaces.skype.com"];
    console.log(`Using this token: ${token}`);

    await fetch("https://presence.teams.microsoft.com/v1/me/forceavailability/", {
      "headers": {
        "accept": "json",
        "accept-language": "en-US",
        "authorization": `Bearer ${token}`,
        "behavioroverride": "redirectAs404",
        "content-type": "application/json",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "x-ms-client-env": "pds-prod-azsc-usea-01",
        "x-ms-client-type": "desktop",
        "x-ms-client-version": "41/1.0.0.2021121106",
        "x-ms-correlation-id": "<get-from-teams>",
        "x-ms-endpoint-id": "<get-from-teams>",
        "x-ms-scenario-id": "336",
        "x-ms-session-id": "<get-from-teams>",
        "x-ms-user-type": "null"
      },
      "referrer": "https://teams.microsoft.com/_",
      "referrerPolicy": "no-referrer-when-downgrade",
      "body": "{\"availability\":\"Available\"}",
      "method": "PUT",
      "mode": "cors"
    }).then(response => {
      let date = new Date();
      console.log(`(PUT) forceavailability(${counter}): status ${response.status} - ${date.toUTCString()}`);
    });

    await fetch("https://presence.teams.microsoft.com/v1/me/reportmyactivity/", {
	  "headers": {
	    "accept": "json",
	    "accept-language": "en-US",
	    "authorization": `Bearer ${token}`,
	    "behavioroverride": "redirectAs404",
	    "content-type": "application/json",
	    "sec-fetch-dest": "empty",
	    "sec-fetch-mode": "cors",
	    "sec-fetch-site": "same-site",
	    "x-ms-client-env": "pds-prod-azsc-usea-01",
	    "x-ms-client-type": "desktop",
	    "x-ms-client-version": "41/1.0.0.2021121106",
	    "x-ms-correlation-id": "<get-from-teams>",
	    "x-ms-endpoint-id": "<get-from-teams>",
	    "x-ms-scenario-id": "343",
	    "x-ms-session-id": "<get-from-teams>",
	    "x-ms-user-type": "null"
	  },
	  "referrer": "https://teams.microsoft.com/_",
	  "referrerPolicy": "no-referrer-when-downgrade",
	  "body": "{\"endpointId\":\"<get-from-teams>","isActive\":true}",
	  "method": "POST",
	  "mode": "cors"
	}).then(response => {
    let date = new Date();
		console.log(`(POST) reportmyactivity(${counter}): status ${response.status} - ${date.toUTCString()}`);
	});

    await fetch("https://presence.teams.microsoft.com/v1/me/presence", {
      "headers": {
        "accept": "json",
        "accept-language": "en-US",
        "authorization": `Bearer ${token}`,
        "behavioroverride": "redirectAs404",
        "content-type": "application/json",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "x-ms-client-env": "pds-prod-azsc-usea-01",
        "x-ms-client-type": "desktop",
        "x-ms-client-version": "41/1.0.0.2021121106",
        "x-ms-correlation-id": "<get-from-teams>",
        "x-ms-endpoint-id": "<get-from-teams>",
        "x-ms-scenario-id": "343",
        "x-ms-session-id": "<get-from-teams>",
        "x-ms-user-type": "null"
      },
      "referrer": "https://teams.microsoft.com/_",
      "referrerPolicy": "no-referrer-when-downgrade",
      "method": "GET",
      "mode": "cors"
    }).then(response => {
      let date = new Date();
      console.log(`(GET) presence(${counter}): status ${response.status} - ${date.toUTCString()}`);
      return response.json();
    }).then(data => {
      console.log(`Showing activity as: ${data.activity} & availability as: ${data.availability}`);
    });

    counter++;
}

var myInterval = setInterval(msTeams, 60000);