import { NgModule } from "@angular/core";
import { RouterModule, Routes } from "@angular/router";
import { AnotherComponent } from "./another/another.component";

const routes: Routes = [
    {
        path: "another",
        component: AnotherComponent
    }
];

@NgModule({
    imports: [RouterModule.forRoot(routes, { useHash: true, enableTracing: true })],
    exports: [RouterModule]
})
export class AppRoutingModule { }
