'use strict'
const assert = require('assert')
const util = require('util')

const Writable = require('readable-stream').Writable
const RTFGroup = require('./rtf-group.js')
const RTFParagraph = require('./rtf-paragraph.js')
const RTFSpan = require('./rtf-span.js')
const iconv = require('iconv-lite')

const availableCP = [
  437, 737, 775, 850, 852, 853, 855, 857, 858, 860, 861, 863, 865, 866,
  869, 932, 1125, 1250, 1251, 1252, 1253, 1254, 1257 ]
const codeToCP = {
  0: 'ASCII',
  77: 'MacRoman',
  128: 'SHIFT_JIS',
  129: 'CP949', // Hangul
  130: 'JOHAB',
  134: 'CP936', // GB2312 simplified chinese
  136: 'BIG5',
  161: 'CP1253', // greek
  162: 'CP1254', // turkish
  163: 'CP1258', // vietnamese
  177: 'CP862', // hebrew
  178: 'CP1256', // arabic
  186: 'CP1257',  // baltic
  204: 'CP1251', // russian
  222: 'CP874', // thai
  238: 'CP238', // eastern european
  254: 'CP437' // PC-437
}

class RTFInterpreter extends Writable {
  constructor (document) {
    super({objectMode: true})
    this.doc = document
    this.parserState = this.parseTop
    this.groupStack = []
    this.group = null
    this.once('prefinish', () => this.finisher())
    this.hexStore = []
  }
  _write (cmd, encoding, done) {
    const method = 'cmd$' + cmd.type.replace(/-(.)/g, (_, char) => char.toUpperCase())
    if (this[method]) {
      this[method](cmd)
    } else {
      process.emit('error', `Unknown RTF command ${cmd.type}, tried ${method}`)
    }
    done()
  }
  finisher () {
    while (this.groupStack.length) this.cmd$groupEnd()
    const initialStyle = this.doc.content.length ? this.doc.content[0].style : []
    for (let prop of Object.keys(this.doc.style)) {
      let match = true
      for (let para of this.doc.content) {
        if (initialStyle[prop] !== para.style[prop]) {
          match = false
          break
        }
      }
      if (match) this.doc.style[prop] = initialStyle[prop]
    }
  }
  flushHexStore () {
    if (this.hexStore.length > 0) {
      let hexstr = this.hexStore.map(cmd => cmd.value).join('')
      this.group.addContent(new RTFSpan({
        value: iconv.decode(
          Buffer.from(hexstr, 'hex'), this.group.get('charset'))
      }))
      this.hexStore.splice(0)
    }
  }

  cmd$groupStart () {
    this.flushHexStore()
    if (this.group) this.groupStack.push(this.group)
    this.group = new RTFGroup(this.group || this.doc)
  }
  cmd$ignorable () {
    this.flushHexStore()
    this.group.ignorable = true
  }
  cmd$endParagraph () {
    this.flushHexStore()
    this.group.addContent(new RTFParagraph())
  }
  cmd$groupEnd () {
    this.flushHexStore()
    const endingGroup = this.group
    this.group = this.groupStack.pop()
    const doc = this.group || this.doc
    if (endingGroup instanceof FontTable) {
      doc.fonts = endingGroup.table
    } else if (endingGroup instanceof ColorTable) {
      doc.colors = endingGroup.table
    } else if (endingGroup instanceof ListTable) {
      doc.lists = endingGroup.table
    } else if (endingGroup instanceof ListOverrideTable) {
      doc.overrideLists = endingGroup.table
    } else if (endingGroup !== this.doc && !endingGroup.get('ignorable')) {
      for (const item of endingGroup.content) {
        doc.addContent(item)
      }
      process.emit('debug', 'GROUP END', endingGroup.type, endingGroup.get('ignorable'))
    }
  }
  cmd$text (cmd) {
    this.flushHexStore()
    if (!this.group) { // an RTF fragment, missing the {\rtf1 header
      this.group = this.doc
    }
    this.group.addContent(new RTFSpan(cmd))
  }
  cmd$controlWord (cmd) {
    this.flushHexStore()
    if (!this.group.type) this.group.type = cmd.value
    const method = 'ctrl$' + cmd.value.replace(/-(.)/g, (_, char) => char.toUpperCase())
    if (this[method]) {
      this[method](cmd.param)
    } else {
      if (!this.group.get('ignorable')) process.emit('debug', method, cmd.param)
    }
  }
  cmd$hexchar (cmd) {
    this.hexStore.push(cmd)
  }

  ctrl$rtf () {
    this.group = this.doc
  }

  // new line
  ctrl$line () {
    this.group.addContent(new RTFSpan({ value: '\n' }))
  }

  // alignment
  ctrl$qc () {
    this.group.style.align = 'center'
  }
  ctrl$qj () {
    this.group.style.align = 'justify'
  }
  ctrl$ql () {
    this.group.style.align = 'left'
  }
  ctrl$qr () {
    this.group.style.align = 'right'
  }

  // text direction
  ctrl$rtlch () {
    this.group.style.dir = 'rtl'
  }
  ctrl$ltrch () {
    this.group.style.dir = 'ltr'
  }

  // general style
  ctrl$par () {
    this.group.addContent(new RTFParagraph())
  }
  ctrl$pard () {
    this.group.resetStyle()
  }
  ctrl$plain () {
    this.group.style.fontSize = this.doc.getStyle('fontSize')
    this.group.style.bold = this.doc.getStyle('bold')
    this.group.style.italic = this.doc.getStyle('italic')
    this.group.style.underline = this.doc.getStyle('underline')
  }
  ctrl$b (set) {
    this.group.style.bold = set !== 0
  }
  ctrl$i (set) {
    this.group.style.italic = set !== 0
  }
  ctrl$u (num) {
    var charBuf = Buffer.alloc ? Buffer.alloc(2) : new Buffer(2)
    // RTF, for reasons, represents unicode characters as signed integers
    // thus managing to match literally no one.
    charBuf.writeInt16LE(num, 0)
    this.group.addContent(new RTFSpan({value: iconv.decode(charBuf, 'ucs2')}))
  }
  ctrl$super () {
    this.group.style.valign = 'super'
  }
  ctrl$sub () {
    this.group.style.valign = 'sub'
  }
  ctrl$nosupersub () {
    this.group.style.valign = 'normal'
  }
  ctrl$strike (set) {
    this.group.style.strikethrough = set !== 0
  }
  ctrl$scaps (set) {
    this.group.style.smallcaps = set !== 0
  }
  ctrl$ul (set) {
    this.group.style.underline = set !== 0
  }
  ctrl$ulnone (set) {
    this.group.style.underline = false
  }
  ctrl$fi (value) {
    this.group.style.firstLineIndent = value
  }
  ctrl$cufi (value) {
    this.group.style.firstLineIndent = value * 100
  }
  ctrl$li (value) {
    this.group.style.indent = value
  }
  ctrl$lin (value) {
    this.group.style.indent = value
  }
  ctrl$culi (value) {
    this.group.style.indent = value * 100
  }

// encodings
  ctrl$ansi () {
    this.group.charset = 'ASCII'
  }
  ctrl$mac () {
    this.group.charset = 'MacRoman'
  }
  ctrl$pc () {
    this.group.charset = 'CP437'
  }
  ctrl$pca () {
    this.group.charset = 'CP850'
  }
  ctrl$ansicpg (codepage) {
    if (availableCP.indexOf(codepage) === -1) {
      this.emit('error', new Error('Codepage ' + codepage + ' is not available.'))
    } else {
      this.group.charset = 'CP' + codepage
    }
  }

// fonts
  ctrl$fonttbl () {
    this.group = new FontTable(this.group.parent)
  }
  ctrl$f (num) {
    if (this.group instanceof FontTable) {
      this.group.currentFont = this.group.table[num] = new Font()
    } else if (this.group.parent instanceof FontTable) {
      this.group.parent.currentFont = this.group.parent.table[num] = new Font()
    } else {
      this.group.style.font = num
    }
  }
  ctrl$fnil () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'nil'
    }
  }
  ctrl$froman () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'roman'
    }
  }
  ctrl$fswiss () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'swiss'
    }
  }
  ctrl$fmodern () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'modern'
    }
  }
  ctrl$fscript () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'script'
    }
  }
  ctrl$fdecor () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'decor'
    }
  }
  ctrl$ftech () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'tech'
    }
  }
  ctrl$fbidi () {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'bidi'
    }
  }
  ctrl$fcharset (code) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      let charset = null
      if (code === 1) {
        charset = this.group.get('charset')
      } else {
        charset = codeToCP[code]
      }
      if (charset == null) {
        return this.emit('error', new Error('Unsupported charset code #' + code))
      }
      this.group.get('currentFont').charset = charset
    }
  }
  ctrl$fprq (pitch) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').pitch = pitch
    }
  }

  // colors
  ctrl$colortbl () {
    this.group = new ColorTable(this.group.parent)
  }
  ctrl$red (value) {
    if (this.group instanceof ColorTable) {
      this.group.red = value
    }
  }
  ctrl$blue (value) {
    if (this.group instanceof ColorTable) {
      this.group.blue = value
    }
  }
  ctrl$green (value) {
    if (this.group instanceof ColorTable) {
      this.group.green = value
    }
  }
  ctrl$cf (value) {
    this.group.style.foreground = value
  }
  ctrl$cb (value) {
    this.group.style.background = value
  }
  ctrl$fs (value) {
    this.group.style.fontSize = value
  }

  //lists
  ctrl$listtable () {
    this.group = new ListTable(this.group.parent)
  }
  ctrl$list () {
    this.group = new List(this.group.parent)
    this.group.parent.table.push(this.group.list)
  }
  ctrl$listlevel () {
    this.group = new ListLevel(this.group.parent)
    this.group.parent.list.levels.push(this.group.listlevel)
  }

  ctrl$listid (num) {
    if (this.group instanceof List || ListOverride) {
      this.group.list.id = num
    }
  }
  ctrl$listtemplateid (num) {
    if (this.group instanceof List) {
      this.group.list.templateid = num
    }  
  }
  ctrl$listsimple (num) {
    if (this.group instanceof List) {
      this.group.list.simple = num
    }
  }
  ctrl$listhybrid () {
    if (this.group instanceof List) {
      this.group.list.hybrid = true
    }
  }
  ctrl$listname (text) {
    if (this.group instanceof List) {
      this.group.list.name = text
    }
  }
  ctrl$liststyleid (num) {
    if (this.group instanceof List) {
      this.group.list.styleid = num
    }
  }
  ctrl$liststylename (txt) {
    if (this.group instanceof List) {
      this.group.list.stylename = text
    }
  }
  ctrl$liststartat (num) {
    if (this.group instanceof List) {
      this.group.list.startat = num
    }
  }
  ctrl$lvltentative () {
    if (this.group instanceof List) {
      this.group.list.tentative = true
    }
  }

  ctrl$levelstartat (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.startat = num
    }
  }
  ctrl$levelnfc (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.nfc = num
    }  
  }
  ctrl$levelnfcn (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.nfcn = num
    }
  }
  ctrl$leveljc (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.jc = num
    }
  }
  ctrl$leveljcn (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.jcn = num
    }
  }
  ctrl$leveltext (txt) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.text = txt
    }
  }
  ctrl$levelnumbers (txt) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.numbers = txt
    }
  }
  ctrl$levelfollow (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.follow = num
    }
  }
  ctrl$levellegal (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.legal = num
    }
  }
  ctrl$levelnorestart (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.norestart = num
    }
  }
  ctrl$levelold (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.old = num
    }
  }
  ctrl$levelprev (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.prev = num
    }
  }
  ctrl$levelprevspace (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.prevspace = num
    }
  }
  ctrl$levelindent (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.indent = num
    }
  }
  ctrl$levelspace (num) {
    if (this.group instanceof ListLevel) {
      this.group.listlevel.space = num
    }
  }

  ctrl$ilvl (num) {
    this.group.style.ilvl = num
  }

  //list overrides
  ctrl$listoverridetable () {
    this.group = new ListOverrideTable(this.group.parent)
  }
  ctrl$listoverride () {
    this.group = new ListOverride(this.group.parent)
    this.group.parent.table.push(this.group.list)
  }

  //list overrides
  ctrl$listoverridecount (num) {
    if (this.group instanceof ListOverride) {
      this.group.list.overridecount = num
    }
  }
  ctrl$ls (num) {
    if (this.group instanceof ListOverride) {
      this.group.list.ls = num
    } else {
      this.group.style.ls = num
    }
  }
  ctrl$listoverridestartat () {
    if (this.group instanceof ListOverride) {
      this.group.list.overridestartat = true
    }
  }
  ctrl$listoverrideformat (num) {
    if (this.group instanceof ListOverride) {
      this.group.list.overrideformat = num
    }
  }

// margins
  ctrl$margl (value) {
    this.doc.marginLeft = value
  }
  ctrl$margr (value) {
    this.doc.marginRight = value
  }
  ctrl$margt (value) {
    this.doc.marginTop = value
  }
  ctrl$margb (value) {
    this.doc.marginBottom = value
  }

// unsupported (and we need to ignore content)
  ctrl$stylesheet (value) {
    this.group.ignorable = true
  }
  ctrl$info (value) {
    this.group.ignorable = true
  }
  ctrl$mmathPr (value) {
    this.group.ignorable = true
  }
}

class FontTable extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.table = []
    this.currentFont = {family: 'roman', charset: 'ASCII', name: 'Serif'}
  }
  addContent (text) {
    this.currentFont.name += text.value.replace(/;\s*$/, '')
  }
}

class Font {
  constructor () {
    this.family = null
    this.charset = null
    this.name = ''
    this.pitch = 0
  }
}

class ColorTable extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.table = []
    this.red = 0
    this.blue = 0
    this.green = 0
  }
  addContent (text) {
    assert(text.value === ';', 'got: ' + util.inspect(text))
    this.table.push({
      red: this.red,
      blue: this.blue,
      green: this.green
    })
    this.red = 0
    this.blue = 0
    this.green = 0
  }
}

class ListTable extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.table = []
  }
}
class List extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.list = {
      levels: []
    }
  }
}
class ListLevel extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.listlevel = {
      style: this.style
    }
  }
}

class ListOverrideTable extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.table = []
  }
}
class ListOverride extends RTFGroup {
  constructor (parent) {
    super(parent)
    this.list = {}
  }
}

module.exports = RTFInterpreter
