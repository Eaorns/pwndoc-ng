let docx = require("docx");
let xml = require("xml");
let htmlparser = require("htmlparser2");


let tableDefaults = {
  "border": {
      "_attr": {
        "w:val": "single",
        "w:color": "FFFFFF",
        "w:sz": 24,
        "w:themeColor": "background1",
        "w:space": 0
    }
  },
  "w:rPr": {
    "header": {
      "w:rPr": [{
        "w:color": {
          "_attr": {
            "w:val": "FFFFFF",
            "w:themeColor": "background1"
          }
        }
      },
      {
        "w:b": {
          "_attr": {
            "w:val": "true"
          }
        }
      }]
    }
  },
  "w:pPr": [
    {
      // "w:pStyle": {
      //   "_attr": {
      //     "w:val": "Normal-Linespacing"
      //   }
      // },
      "w:spacing": {
        "_attr": {
          "w:after": 0,
          "w:line": 276,
          "w:lineRule": "auto"
        }
      }
    }
  ],
  "w:shd": {
    "header": {
      "w:shd": {
        "_attr": {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": "2DA3DE"
        }
      }
    },
    "rowu": {
      "w:shd": {
        "_attr": {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": "D9D9D9"
        }
      }
    },
    "rowe": {
      "w:shd": {
        "_attr": {
          "w:val": "clear",
          "w:color": "auto",
          "w:fill": "F2F2F2"
        }
      }
    }
  },
  "w:gridSpan": {
    "w:gridSpan": {
      "_attr": {
        "w:val": 1
      }
    }
  }
};


function html2ooxml(html, style = "") {
  if (html === "") return html;
  if (!html.match(/^<.+>/)) html = `<p>${html}</p>`;
  let doc = new docx.Document({ sections: [] });
  let paragraphs = [];
  let cParagraph = null;
  let cRunProperties = {};
  let cParagraphProperties = {};
  let list_state = [];
  let inCodeBlock = false;
  let inTable = false;
  let inTableRow = false;
  let inTableCell = false;
  let cellHasText = false;
  let tmpAttribs = {};
  let tableHeader = false
  let tmpTable = [];
  let tmpCells = [];
  let tmpCellContent = [];
  let parser = new htmlparser.Parser(
    {
      onopentag(tag, attribs) {
        switch (tag) {
          case 'h1':
            cParagraph = new docx.Paragraph({ heading: "Heading1" });
            break;
          case 'h2':
            cParagraph = new docx.Paragraph({ heading: "Heading2" });
            break;
          case 'h3':
            cParagraph = new docx.Paragraph({ heading: "Heading3" });
            break;
          case 'h4':
            cParagraph = new docx.Paragraph({ heading: "Heading4" });
            break;
          case 'h5':
            cParagraph = new docx.Paragraph({ heading: "Heading5" });
            break;
          case 'h6':
            cParagraph = new docx.Paragraph({ heading: "Heading6" });
            break;
          case 'div':
          case 'p':
            // Do not overwrite style if style is already set, for example with bullet points.
            if (style && typeof style === 'string' && cParagraphProperties.style == undefined) {
              cParagraphProperties.style = style
            }
            cParagraph = new docx.Paragraph(cParagraphProperties)
            break;
          case 'table':
            inTable = true;
            break;
          case 'td':
            tmpAttribs = attribs;
            inTableCell = true;
            cellHasText = false;
            tmpCellContent = [];
            break;
          case 'th':
            inTableCell = true;
            tableHeader = true;
            tmpAttribs = attribs;
            tmpCellContent = [];
            cellHasText = false;
            break;
          case 'tr':
            inTableRow = true;
            break;
          case 'pre':
            inCodeBlock = true;
            cParagraph = new docx.Paragraph({ style: "Code" });
            break;
          case 'br':
            cParagraph.addChildElement(new docx.Run({ break: 1 }))
            break;
          case 'b':
          case 'strong':
            cRunProperties.bold = true;
            break;
          case 'i':
          case 'em':
            cRunProperties.italics = true;
            break;
          case 'u':
            cRunProperties.underline = {};
            break;
          case 'strike':
          case 's':
            cRunProperties.strike = true;
            break;
          case 'mark':
            //Possible values are: black, blue, cyan, darkBlue, darkCyan, darkGray, darkGreen, darkMagenta, darkRed, darkYellow, green, lightGray, magenta, none, red, white, yellow
            let color;
            switch (attribs["data-color"]) {
              case "#ffff00":
                color = "yellow";
                break;
              case "#fe0000":
                color = "red";
                break;
              case "#00ff00":
                color = "green";
                break;
              case "#00ffff":
                color = "cyan";
                break;
            }
            cRunProperties.highlight = color;
            break;
          case 'a':
            cRunProperties.link = attribs.href;
            break;
          case 'ul':
            // list_state.push("bullet");
            list_state.push("number");  // NOTE swapped to fix weird bug where numbered lists became bullets and vice versa
            break;
          case 'ol':
            // list_state.push("number");
            list_state.push("bullet");  // NOTE swapped to fix weird bug where numbered lists became bullets and vice versa
            break;
          case 'li':
            let level = list_state.length - 1;
            if (level >= 0 && list_state[level] === "bullet")
              cParagraphProperties.bullet = { level: level };
            else if (level >= 0 && list_state[level] === "number")
              cParagraphProperties.numbering = { reference: 2, level: level };
            else cParagraphProperties.bullet = { level: 0 };
            cParagraphProperties.style = (list_state[level] == "bullet") ? "numberStyle" : "bulletStyle";  // NOTE also the weird number/bullet switching bug...
            break;
          case 'code':
            cRunProperties.style = "CodeChar";
            break;
          case 'legend':
            if (attribs && attribs.alt !== "undefined") {
              let label = attribs.label || "Figure";
              cParagraph = new docx.Paragraph({
                style: "Caption",
                alignment: docx.AlignmentType.CENTER,
              });
              cParagraph.addChildElement(new docx.TextRun(`${label} `));
              cParagraph.addChildElement(new docx.SimpleField(`SEQ ${label}`, "1"));
              cParagraph.addChildElement(new docx.TextRun(` - ${attribs.alt}`));
            }
            break;
        }
      },

      ontext(text) {
        if (cRunProperties.link) {
          cParagraph.addChildElement(new docx.TextRun({ "text": `{_|link|_{${text}|-|${cRunProperties.link}}_|link|_}`, "style": "Hyperlink" }));

        } else if (text && cParagraph) {
          if (inTableCell) {
            cellHasText = true;
          }
          cRunProperties.text = text;
          cParagraph.addChildElement(new docx.TextRun(cRunProperties));
        }
      },

      onclosetag(tag) {
        if (
          [
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
            "div",
            "p",
            "pre",
            "img",
            "legend",
            //"table",
            /* "tr",
            "th", */
          ].includes(tag)) {

          if (inTableCell) {
            tmpCellContent.push(cParagraph)
          } else {
            paragraphs.push(cParagraph);
          }

          cParagraph = null;
          cParagraphProperties = {};
          if (tag === "pre") inCodeBlock = false;
        } else if (tag === "b" || tag === "strong") {
          delete cRunProperties.bold;
        } else if (tag === "i" || tag === "em") {
          delete cRunProperties.italics;
        } else if (tag === "u") {
          delete cRunProperties.underline;
        } else if (tag === "mark") {
          delete cRunProperties.highlight;
        } else if (tag === "strike" || tag === "s") {
          delete cRunProperties.strike;
        } else if (tag === "ul" || tag === "ol") {
          list_state.pop();
          if (list_state.length === 0) cParagraphProperties = {};
        } else if (tag === "code") {
          delete cRunProperties.style;
        } else if (tag === "tr") {
          inTableRow = false;
          tableHeader = false;
          tmpTable.push(tmpCells);
          tmpCells = []
        } else if (tag === "a") {
          delete cRunProperties.link
        } else if (tag === "td" || tag === "th") {
          tmpCells.push({
            text: cellHasText === true ? tmpCellContent : "",
            width: tmpAttribs.colwidth ? tmpAttribs.colwidth : "250",
            header: tableHeader,
          });

          tmpAttribs = {};
          tmpCellContent = [];
          inTableCell = false;
        } else if (tag === "table") {
          inTable = false;
          let tblRows = [];
          tmpTable.map((row) => {
            let tmpCells = [];
            let isHeader = false
            let widthTotal = row.map(cell => parseInt(cell.width)).reduce((prev, next) => prev + next);

            row.map((cell) => {
              isHeader = cell.header;
              tmpCells.push(new docx.TableCell({
                width: {
                  size: Math.round(cell.width / widthTotal),
                  type: "pct",
                },
                children: cell.text,
              }))
            });

            tblRows.push(new docx.TableRow({
              children: tmpCells,
              tableHeader: isHeader,
            }))


          });
          // build table and push to paragraphs array
          cParagraph = new docx.Table({
            rows: tblRows,
            width: {
              size: 100,
              type: "pct"
            }
          });

          paragraphs.push(cParagraph);
          cParagraph = null;
          cParagraphProperties = {};
          tmpTable = [];
          tmpCells = [];
        }
      },

      onend() {
        doc.addSection({
          children: paragraphs,
        });
      },
    },
    { decodeEntities: true }
  );

  // For multiline code blocks
  html = html.replace(/\n/g, "<br>");
  parser.write(html);
  parser.end();

  let prepXml = doc.documentWrapper.document.body.prepForXml({});
  let filteredXml = prepXml["w:body"].filter((e) => {
    return Object.keys(e)[0] === "w:p" || Object.keys(e)[0] === "w:tbl";
  });

  let traverseTable = (obj, inTable, isHeader) => {
    if (typeof obj != "object")  // exit condition
      return obj;

    Object.keys(obj).map((key, idx) => {
      if (!inTable) {
        obj[key] = traverseTable(obj[key], key == "w:tbl", isHeader);
        return obj;
      }
      switch (key) {
        case "w:tblPr":
            obj[key] = traverseTable(obj[key], inTable, isHeader);
            obj[key].unshift({"w:tblLayout": {"_attr": {"w:type": "fixed"}}});
            obj[key].unshift({"w:tblStyle": {"_attr": {"w:val": "TableGrid"}}});
            break;
        case "w:tblBorders":
          obj[key] = [
            {"w:top": tableDefaults['border']},
            {"w:left": tableDefaults['border']},
            {"w:bottom": tableDefaults['border']},
            {"w:right": tableDefaults['border']},
            {"w:insideH": tableDefaults['border']},
            {"w:insideV": tableDefaults['border']},
          ];
          break;
        // case "w:tblGrid":
        //   obj[key] = Array(obj[key].length).fill({"w:gridCol": {"_attr": {"w:w": Math.round(5100 / obj[key].length)}}});
        //   break;
        case "w:tr":
          isHeader = obj[key].map((e) => ("w:trPr" in e && (Object.keys(e["w:trPr"][0]["w:tblHeader"]).length == 0 ||
                                                            ("_attr" in e["w:trPr"][0]["w:tblHeader"] &&
                                                             "w:val" in e["w:trPr"][0]["w:tblHeader"]["_attr"] &&
                                                             e["w:trPr"][0]["w:tblHeader"]["_attr"]["w:val"])))
                                 ).some((e) => e);
          rowNumber = (isHeader) ? 0 : rowNumber + 1;
          obj[key] = traverseTable(obj[key], inTable, isHeader);
          break;
        case "w:tcPr":
          obj[key] = traverseTable(obj[key], inTable, isHeader)
          obj[key].unshift(tableDefaults['w:shd'][isHeader ? 'header' : (rowNumber % 2 == 0 ? 'rowe' : 'rowu')]);
          obj[key].unshift(tableDefaults['w:gridSpan']);
          break;
        case "w:pPr":
          obj[key] = tableDefaults['w:pPr'];
          break;
        case "w:r":
          obj[key] = traverseTable(obj[key], inTable, isHeader)
          if (isHeader)
            obj[key].unshift(tableDefaults['w:rPr']['header']);
          break;
        default:
          obj[key] = traverseTable(obj[key], inTable, isHeader);
      }
    });
    return obj;
  }
  let rowNumber = 0;
  traverseTable(filteredXml, false, false);

  let dataXml = xml(filteredXml);
  dataXml = dataXml.replace(/w:numId w:val="{2-0}"/g, 'w:numId w:val="2"'); // Replace numbering to have correct value
  //a little dirty but until we do better it works
  dataXml = dataXml.replace(/\{_\|link\|_\{(.*?)\|\-\|(.*?)\}_\|link\|_\}/gm, '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK $2 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t> $1 </w:t> </w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
  return dataXml;
}

module.exports = html2ooxml;
