/* eslint-disable prettier/prettier */
/* eslint-disable no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// import "zone.js"; // Required for Angular
import { platformBrowserDynamic } from "@angular/platform-browser-platformBrowserDynamic";
import AppModule from "./app/app.module";
/* global console, Office */

/* (async () => {
  console.log("taskpane 1");
  await Office.onReady();
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch((error) => console.error(error));
})();

Office.initialize = () => {}; */
console.log("taskpane 1");
// platformBrowserDynamic().bootstrapModule(AppModule);
(<any>window).Office.initialize = (reason) => {
  console.log("Initializing Office");
  console.log("Bootstraping module");
  platformBrowserDynamic().bootstrapModule(AppModule);

  // Set up ItemChanged event
  // (<any>window).Office.context.mailbox.addHandlerAsync((<any>window).Office.EventType.ItemChanged, itemChanged);
};
