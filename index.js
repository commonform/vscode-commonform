const fs = require('fs')
const path = require('path')
const vscode = require('vscode')

const commonmark = require('commonform-commonmark')
const docx = require('commonform-docx')
const numberings = {
  outline: require('outline-numbering'),
  decimal: require('decimal-numbering')
}
const prepareBlanks = require('commonform-prepare-blanks')
const signaturePages = require('ooxml-signature-pages')
const html = require('commonform-html')

const outputChannel = vscode.window.createOutputChannel('Common Form')

exports.activate = context => {
  console.log('Loaded vscode-commonform.')

  const docxCommand = vscode.commands.registerCommand('vscode-commonform.docx', () => {
    const document = vscode.window.activeTextEditor.document
    const fileName = path.normalize(document.fileName)
    const dirname = path.dirname(fileName)
    const extname = path.extname(fileName)
    const basename = path.basename(fileName, extname)
    const outputPath = path.join(dirname, `${basename}.docx`)

    let parsed
    try {
      parsed = commonmark.parse(document.getText())
    } catch (error) {
      return fail('Error parsing markup.')
    }

    const { form, frontMatter, directions } = parsed

    const options = {
      numbering: numberings.outline,
      title: frontMatter.title || 'Untitled Form',
      edition: frontMatter.edition || undefined,
      hash: !!frontMatter.hash,
      leftAlignBody: Boolean(frontMatter.leftAlignBody),
      a4: Boolean(frontMatter.a4),
      markFilled: Boolean(frontMatter.markFilled),
      smartify: frontMatter.smartify === undefined ? true : !!frontMatter.smartify,
      centerTitle: !!frontMatter.centerTitle,
      indentMargins: !!frontMatter.indentMargins,
      styles: frontMatter.styles
    }

    if (frontMatter.numbering) {
      if (typeof frontMatter.numbering === 'string') {
        const specified = numberings[frontMatter.numbering]
        if (specified) {
          options.numbering = specified
        } else {
          return fail(`No such numbering scheme: ${frontMatter.numbering}`)
        }
      } else {
        return fail('Numbering is not a string.')
      }
    }

    if (frontMatter.signatures) {
      options.after = signaturePages(frontMatter.signatures)
    }

    let blanks
    if (frontMatter.blanks) {
      if (Array.isArray(frontMatter.blanks)) {
        blanks = frontMatter.blanks
      } else {
        blanks = prepareBlanks(frontMatter.blanks, directions)
      }
    } else {
      blanks = []
    }

    const writeStream = fs.createWriteStream(outputPath)
    docx(form, blanks, options)
      .generateNodeStream()
      .pipe(writeStream)
      .once('end', () => {
        const message = `Wrote ${outputPath}`
        console.log(message)
        outputChannel.append(message + '\n')
      })

    function fail (message) {
      vscode.window.showErrorMessage(`stderr: ${message}\n`)
      outputChannel.append(`stderr: ${message}\n`)
    }
  })

  context.subscriptions.push(docxCommand)

  const htmlCommand = vscode.commands.registerCommand('vscode-commonform.html', () => {
    const document = vscode.window.activeTextEditor.document
    const fileName = path.normalize(document.fileName)
    const dirname = path.dirname(fileName)
    const extname = path.extname(fileName)
    const basename = path.basename(fileName, extname)
    const outputPath = path.join(dirname, `${basename}.html`)

    let parsed
    try {
      parsed = commonmark.parse(document.getText())
    } catch (error) {
      return fail('Error parsing markup.')
    }

    const { form, frontMatter, directions } = parsed

    const options = {
      title: frontMatter.title || 'Untitled Form',
      edition: frontMatter.edition || undefined,
      depth: frontMatter.edition || undefined,
      classNames: frontMatter.classNames || undefined,
      html5: true,
      ids: true,
      lists: true
    }

    let blanks
    if (frontMatter.blanks) {
      if (Array.isArray(frontMatter.blanks)) {
        blanks = frontMatter.blanks
      } else {
        blanks = prepareBlanks(frontMatter.blanks, directions)
      }
    } else {
      blanks = []
    }

    let output
    try {
      output = html(form, blanks, options)
    } catch (error) {
      return fail(error.toString())
    }

    fs.writeFile(outputPath, output, error => {
      if (error) return fail(error.toString())
      const message = `Wrote ${outputPath}`
      console.log(message)
      outputChannel.append(message + '\n')
    })

    function fail (message) {
      vscode.window.showErrorMessage(`stderr: ${message}\n`)
      outputChannel.append(`stderr: ${message}\n`)
    }
  })

  context.subscriptions.push(htmlCommand)
}

exports.deactivate = () => {}
