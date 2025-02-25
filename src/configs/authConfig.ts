import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
    auth: {
        clientId: process.env.WEBPACK_CLIENT_ID,
        authority: "https://login.microsoftonline.com/273f45e0-e235-4dde-ab7a-fd3e631a88e0",
        redirectUri: `/taskpane.html`,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true,
    },
}

const msalInstance = new PublicClientApplication(msalConfig);

export { msalInstance };