import { Component, ContentChild } from "@angular/core";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  formattedText: string
  SQL_CORE_KEYWORDS: string[] = ['SELECT', 'FROM', 'WHERE', 'AND', 'IN']

  async newCodeBlockOutline() {
    try {
      await OneNote.run(async context => {
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

  formatSQLText(text: string): string {
    let formattedText: string = ''
    let words = text.split(' ')
    if (words.length != 0) {
      words.forEach(word => {
        word = word.toUpperCase()
        this.SQL_CORE_KEYWORDS.includes(word) ?
        formattedText += '<span style="font-weight: bold; color: blue"> ' + word + '</span>':
        formattedText += '<span> '+ word + '<span>'
      })
    }
    return formattedText
  }

  async writeHTMLToPage(HTML: string) {
    try {
      await OneNote.run(async context => {
        let page = context.application.getActivePage()
        page.addOutline(100, 100,
          HTML
        )
        return context.sync()
      })
    } catch (e) {
      console.log(e)
    }
  }

  async processCodeBlock(codestyle: string) {
    try {
      await OneNote.run(async context => {
        let paragraphs = context.application.getActiveOutline().paragraphs
        paragraphs.load('richText')
        context.sync().then(() => {
          paragraphs.items.forEach(item => {
            let HTML = this.formatSQLText(item.richText.text)
            this.writeHTMLToPage(HTML)
          })
        })
      })
    } catch (e) {
      console.log(e)
    }
  }


}
