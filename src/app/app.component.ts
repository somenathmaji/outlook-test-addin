import { Component, OnInit } from "@angular/core";
import { Router } from "@angular/router";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit{

  constructor(
    private router: Router
  ) {}

  ngOnInit(): void {
    (<any>window).Office.context.mailbox.addHandlerAsync(
      (<any>window).Office.EventType.ItemChanged, 
      () => {console.error('hello!!'); this.itemChangedEventHandler()}, 
      () => {console.error('registered')});
  }

  public itemChangedEventHandler = () => {
    this.switchView();
  }   

  public switchView = () => {
    this.router.navigate(["another"], { replaceUrl: true });
  }
}
