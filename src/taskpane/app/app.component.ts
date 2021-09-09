/* eslint-disable prettier/prettier */
/* eslint-disable no-prototype-builtins */
/* eslint-disable no-undef */
import { Component, OnInit } from "@angular/core";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template,
})
export default class AppComponent implements OnInit {
  subject: string = "empty";

  // eslint-disable-next-line no-unused-vars
  constructor() {
    console.log("constructor");
  }

  ngOnInit() {
    console.log("ngOnInit app component");
    if (this.isItemChangedEventAvailable()) {
      console.log("itemchanged available");
      (<any>window).Office.context.mailbox.addHandlerAsync((<any>window).Office.EventType.ItemChanged, () => this.itemChanged());
    } else {
      console.log("itemchanged not available");
    }
  }

  private isItemChangedEventAvailable() {
    const eventItemChanged = "ItemChanged";
    if (!window) {
      console.warn("Window object not detected");
      return false;
    }
    const Office = (<any>window).Office;
    if (!Office) {
      console.warn("Office object not detected");
      return false;
    }
    if (!Office.EventType) {
      console.warn("EventType object not detected");
      return false;
    }
    if (!Office.EventType.hasOwnProperty) {
      console.warn("EventType object does not have haveOwnProperty");
      return false;
    }
    if (!Office.EventType.hasOwnProperty(eventItemChanged)) {
      console.warn("ItemChanged not detected");
      return false;
    }
    return true;
  }

  private itemChanged() {
    console.log("itemchanged event occured");
    var item = (<any>window).Office.context.mailbox.item;
    if (item) {
      console.log("item available");
      this.subject = item.subject;
      console.log(this.subject);
    } else {
      console.log("item is null/undefined");
      this.subject = "couldn't fetch current item";
      console.log(this.subject);
    }
  }
}
