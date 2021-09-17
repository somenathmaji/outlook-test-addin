import { ChangeDetectorRef, Component, NgZone, OnInit } from "@angular/core";
import { OfficeService } from "./office.service";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit{
  
  public subject = "";

  constructor(
    private officeService : OfficeService
  ) {}

  ngOnInit(): void {
    this.subject = (window as any).Office.context.mailbox.item.subject;
    if (this.officeService.isItemChangedEventAvailable()) {
      console.log("[app-component] Registering to ItemChanged event");
      this.officeService.registerEventHandler((window as any).Office.EventType.ItemChanged, func => this.updateSubject());
    } else {
      console.warn("[app-component] Skipping registering for ItemChanged event");
    }
  }

  public updateSubject() {
    this.subject = (window as any).Office.context.mailbox.item.subject;
    console.log(this.subject);
  }
}
