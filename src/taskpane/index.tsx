import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { msalInstance } from "../configs/authConfig";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const getAccessToken = async () => {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("No account logged in");
    }

    const response = await msalInstance.acquireTokenSilent({
      scopes: ["Files.Read", "Files.Read.All"],
      account: accounts[0],
    })

    localStorage.setItem("token", response.accessToken);
  } catch (error) {
    console.log("Error getting token -> ", error);
    localStorage.removeItem("token");
  }
}

const loginUser = async () => {
  try {
    await msalInstance.initialize();
    await msalInstance.loginPopup({
      scopes: ["User.read", "Files.Read", "Files.Read.All", "Sites.Read.All"],
      prompt: "select_account",
    });
    await getAccessToken();
  } catch (error) {
    console.log("Login failed -> ", error);
    throw Error(`Login Failed - ${error.message}`);
  }
}

/* Render application after Office initializes */
Office.onReady(async () => {
  try {
    await loginUser();
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    );
  } catch (error) {
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <div>
          <h2>Something went wrong</h2>
          <p>{error.message}</p>
        </div>
      </FluentProvider>
    );
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
