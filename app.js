// JavaScript code here
function onSignin(idToken) {
    let user = clientApplication.getAccount();
    document.getElementById("userName").innerHTML = "Currently logged in as " + user.name;
    let requestObj1 = {
      scopes: ["user.read", 'openid', 'profile']
    };
  }

  function onSignInClick() {
    let requestObj = {
      scopes: ["user.read", 'openid', 'profile']
    };

    clientApplication.loginPopup(requestObj).then(onSignin).catch(function (error) { console.log(error) });
  }

  function getOAuthCardResourceUri(activity) {
    if (activity &&
      activity.attachments &&
      activity.attachments[0] &&
      activity.attachments[0].contentType === 'application/vnd.microsoft.card.oauth' &&
      activity.attachments[0].content.tokenExchangeResource) {
      // asking for token exchange with AAD
      return activity.attachments[0].content.tokenExchangeResource.uri;
    }
  }

  function exchangeTokenAsync(resourceUri) {
    let user = clientApplication.getAccount();
    if (user) {
      let requestObj = {
        scopes: [resourceUri]
      };
      return clientApplication.acquireTokenSilent(requestObj)
        .then(function (tokenResponse) {
          return tokenResponse.accessToken;
        })
        .catch(function (error) {
          console.log(error);
        });
    } else {
      return Promise.resolve(null);
    }
  }

  async function fetchJSON(url, options = {}) {
    const res = await fetch(url, {
      ...options,
      headers: {
        ...options.headers,
        accept: 'application/json'
      }
    });

    if (!res.ok) {
      throw new Error(`Failed to fetch JSON due to ${res.status}`);
    }

    return await res.json();
  }

  var clientApplication;
  (function () {
    var msalConfig = {
      auth: {
        clientId: '<CANVAS CLIENT APP ID>',
        authority: 'https://login.microsoftonline.com/<TENANT ID>'
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
      }
    };
    if (!clientApplication) {
      clientApplication = new Msal.UserAgentApplication(msalConfig);
    }
  }());

  (async function main() {
    // Add your BOT ID below 
    var theURL = "<Token endpoint URL>" // you can find the token URL via the mobile app channel configuration

    var userId = clientApplication.account?.accountIdentifier != null ?
      ("You-customized-prefix" + clientApplication.account.accountIdentifier).substr(0, 64) :
      (Math.random().toString() + Date.now().toString()).substr(0, 64);

    const { token } = await fetchJSON(theURL);
    const directLine = window.WebChat.createDirectLine({ token });
    const store = WebChat.createStore({}, ({ dispatch }) => next => action => {
      const { type } = action;
      if (action.type === 'DIRECT_LINE/CONNECT_FULFILLED') {
        dispatch({
          type: 'WEB_CHAT/SEND_EVENT',
          payload: {
            name: 'startConversation',
            type: 'event',
            value: { text: "hello" }
          }
        });
        return next(action);
      }
      if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
        const activity = action.payload.activity;
        let resourceUri;
        if (activity.from && activity.from.role === 'bot' &&
          (resourceUri = getOAuthCardResourceUri(activity))) {
          exchangeTokenAsync(resourceUri).then(function (token) {
            if (token) {
              directLine.postActivity({
                type: 'invoke',
                name: 'signin/tokenExchange',
                value: {
                  id: activity.attachments[0].content.tokenExchangeResource.id,
                  connectionName: activity.attachments[0].content.connectionName,
                  token
                },
                "from": {
                  id: userId,
                  name: clientApplication.account.name,
                  role: "user"
                }
              }).subscribe(
                id => {
                  if (id === 'retry') {
                    // bot was not able to handle the invoke, so display the oauthCard
                    return next(action);
                  }
                  // else: tokenexchange successful and we do not display the oauthCard
                },
                error => {
                  // an error occurred to display the oauthCard
                  return next(action);
                }
              );
              return;
            } else
              return next(action);
          });
        } else
          return next(action);
      } else
        return next(action);
    });

    const styleOptions = {
      // Add styleOptions to customize Web Chat canvas
      hideUploadButton: true
    };

    window.WebChat.renderWebChat(
      {
        directLine: directLine,
        store,
        userID: userId,
        styleOptions
      },
      document.getElementById('webchat')
    );
  })().catch(err => console.error("An error occurred: " + err));
