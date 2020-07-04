import { Component } from "@angular/core";
import { Formatting } from "./formatting.interface";
const template = require("./app.component.html");
/* global require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent implements Formatting {

  async processCodeBlock(codestyle: string) {
    try {
      await OneNote.run(async context => {
        let paragraphs = context.application.getActiveOutline().paragraphs
        paragraphs.load('richText')
        context.sync().then(() => {
          let text = ""
          paragraphs.items.forEach(item => {
            text += item.richText.text + ' \n '
          })
          let HTML = this.getFormatProcess(codestyle)(text) //function currying...?
          if (HTML.length != 0) {
            this.writeHTMLToPage(HTML)
            //this.deleteOldOutline()
          }
        })
      })
    } catch (e) {
      console.log(e)
    }
  }

  getFormatProcess(codestyle: string) {
    switch (codestyle) {
      case 'SQL': return this.formatSQLText
      case 'JS': return this.formatJSText
    }
  }

  formatJSText(text: string): string {
    return '<p>' + text + '</p>'
  }

  formatSQLText(text: string): string {
    const SQL_CORE_KEYWORDS: string[] = ['SELECT', 'FROM', 'WHERE', 'AND', 'IN']
    let formattedText: string = ''
    let words = text.split(' ')
    if (words.length != 0) {
      words.forEach(word => {
        SQL_CORE_KEYWORDS.includes(word.toUpperCase())
          ? formattedText += `<span style="font-weight: bold; color: blue; font-family: 'Consolas', 'monaco', 'monospace'">` + word.toUpperCase() + ' </span>'
          : word == '\n'
            ? formattedText += '<p></p>'
            : formattedText += '<span> ' + word + ' <span>'
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

  // async deleteOldOutline() {
  //   try {
  //     await OneNote.run(async context => {
  //       let outline = context.application.getActiveOutline()
  //       outline.load()
  //       context.sync().then(() => {
  //         outline = null
  //       })
  //     })
  //   } catch (e) {
  //     console.log(e)
  //   }
  //}


}
