import { Component } from "@angular/core";
import { Formatting } from "./formatting.interface";
import { style } from "@angular/core/src/animation/dsl";
const template = require("./app.component.html");
/* global require */

const COMMON_STYLES = `font-family: 'Consolas', 'monaco', 'monospace';`

interface KeywordsStyle {
  keywords: string[],
  style: string
}

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
            this.writeHTMLToPage('<div style="' + COMMON_STYLES + '">' + HTML + '</div>')
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
    const CORE_KEYWORDS = ['function', 'class', 'extends', 'constructor', 'super', 'this', 'let',
      'const', '=>', 'string', 'number', 'boolean']
    const ACTION_KEYWORDS = ['import', 'export', 'if', 'else', 'switch', 'case', 'return', 'default']

    const styles: KeywordsStyle[] = [
      { keywords: CORE_KEYWORDS, style: 'color: blue;' },
      { keywords: ACTION_KEYWORDS, style: 'color: purple;' }
    ]
    return getStyledHTMLText(text, styles, true)
  }

  formatSQLText(text: string): string {
    const CORE_KEYWORDS: string[] = ['SELECT', 'FROM', 'WHERE', 'AND', 'IN']
    const styles: KeywordsStyle[] = [
      { keywords: CORE_KEYWORDS, style: 'color: blue; font-weight: bold' }
    ]
    return getStyledHTMLText(text, styles, false)
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
function getStyledHTMLText(text: string, keywordStyles: KeywordsStyle[], caseSpecific: boolean): string {
  /**TODO: handle extra spaces and special characters*/
  let formattedText: string = ''
  let words = text.split(' ')
  if (words.length != 0) {
    words.forEach(word => {
      if (word == '\n') {
        formattedText += '<p>' + word + ' </p>'
      } else {
        const comparisonWord = caseSpecific ? word : word.toUpperCase()
        let styled: boolean = false
        keywordStyles.forEach(keywordStyle => {
          if (keywordStyle.keywords.includes(comparisonWord)) {
            formattedText += '<span style="' + keywordStyle.style + '">' + word + ' </span>'
            styled = true
          }
        })
        if (!styled) {
          formattedText += '<span>' + word + ' </span>'
        }
      }
    })
  }
  return formattedText
}