import { Injectable } from '@angular/core';

@Injectable()
export class OfficeService {

  private office;

  constructor() { 
    if (!window.hasOwnProperty("Office")) {
      throw new Error("Unable to locate Office API");
    }
    this.office = (<any>window).Office;
  }

  public registerEventHandler(event: any, handlerFunc: Function) {
      try {
        console.log("Registering handler to event: %s", event);
        this.office.context.mailbox.addHandlerAsync(event, handlerFunc);
        console.log("Finished registering handler to event: %s", event);
      } catch (err) {
        console.error("Failed to register event : %s ", event);
      }
  }

  public isItemChangedEventAvailable() {
    const eventItemChanged = "ItemChanged";
    if (!window) {
       console.warn("[office-service] Window object not detected");
       return false;
    }
    const Office = (<any>window).Office;
    if (!Office) {
       console.warn("[office-service] Office object not detected");
       return false;
    }
    if (!Office.EventType) {
       console.warn("[office-service] EventType object not detected");
       return false;
    }
    if (!Office.EventType.hasOwnProperty) {
       console.warn("[office-service] EventType object does not have haveOwnProperty");
       return false;
    }
    if (!Office.EventType.hasOwnProperty(eventItemChanged)) {
       console.warn(`[office-service] ${eventItemChanged} not detected`);
       return false;
    }
    return true;
}

}
