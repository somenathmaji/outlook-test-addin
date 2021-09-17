import { ChangeDetectorRef, Component, NgZone, OnInit } from "@angular/core";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit{
  
  public subject = "";

  ngOnInit(): void {
    this.subject = Office.context.mailbox.item.subject;
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, this.updateSubject);
    console.log(this.subject);
  }

  public updateSubject() {
    this.subject = Office.context.mailbox.item.subject;
    console.log(this.subject);
  }
}
