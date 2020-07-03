import { Component, ContentChild } from "@angular/core";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  code = "'Consolas', 'monaco', 'monospace'"


  async newCodeBlockOutline() {
    try {
      await OneNote.run( async context => {
        let page = context.application.getActivePage()
        page.addOutline(100, 100,
          `<p id="code-block" style="font-family: 'Consolas', 'monaco', 'monospace'">code contents here</p>`
          )
        return context.sync()
      })
    } catch (e) {
      console.log(e)
    }
  }

  async processCodeBlock(style: string){
    try {
      await OneNote.run( async context => {
        let paragraphs = context.application.getActiveOutline().paragraphs
        paragraphs.load('richText')
        context.sync().then(() => {
          paragraphs.items.forEach(item => {
            console.log(item.richText.text)
          })
        })
      })
    } catch (e) {
      console.log(e)
    }
  }
}
