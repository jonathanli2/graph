export const OAuthSettings = {
 appId: '64213508-9802-4e47-93f7-a5f755880ae2',  // wsibtest outlook account

 // appId: '9ef739c6-4524-4613-b4ac-80310ff2aa53',    // haiquan hotmail azure active directory

 // appId: '6371b3e0-41d7-4df8-837d-4997cdedd3ba', //wsib dev active directory
  redirectUri: 'http://localhost:4200',
  scopes: [
    "user.read",
    "mailboxsettings.read",
    "calendars.readwrite"
  ]
};
