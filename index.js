import XLSX from 'xlsx-js-style'
import download from 'downloadjs'


export default function Excelix(fields) {
  let _this = this;
  let headersCount = 0

  if (!fields) {
    console.error('Please, add fields to constructor')
    return
  }

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet([{}])

  XLSX.utils.book_append_sheet(wb, ws, 'sheet')
  
  _this.addRow = (row, { bold, inline=false, r=headersCount, c=0 }={}) => {
    XLSX.utils.sheet_add_aoa(ws, [row], { origin: { r: r, c: c } })

    if (Array.isArray(row)) {
      row.forEach((r, i) => {
        const cell = XLSX.utils.encode_cell({r:headersCount,c: i})
        ws[cell].s = {font: {bold}}
      })
    } else {
      const cellRef = XLSX.utils.encode_cell({r:headersCount,c:c})
      ws[cellRef].s = {font: {bold}}
    }

    headersCount += 1
  }


  _this.addJson = (jsonData) => {
    if (!Array.isArray(jsonData)) {
      console.error('JSON data must be array')
      return
    }
    
    const filteredJsonData = filterJSONObjects(jsonData, fields)  
    const titles = Object.keys(filteredJsonData[0])
    
    // add titles
    const titlesFormatted = []
    titles.forEach(e => {
       titlesFormatted.push(fields[e].text || e)
    })

    _this.addRow(titlesFormatted, { bold: true }) 

    // add json data
    XLSX.utils.sheet_add_json(ws, filteredJsonData,
      { origin:
        {
          r: headersCount,
          c: 0
        },
        skipHeader: true
      })

    headersCount += Object.keys(filteredJsonData).length


    // add summaries

    const sums = [] 

    titles.forEach((c, i) => {
      const isTotal = fields[c].total

      if (!isTotal) {
        sums.push("")
        return
      }

      let sum = 0
      filteredJsonData.forEach(data => {
        if (!isNaN(data[c])) {
          sum += data[c]
        }
      })

      sums.push(sum)
      sums[0] = 'ИТОГО'

    })
    _this.addRow(sums, {bold: true})
  }


  _this.writeToFile = (filename) => {
    const data = XLSX.write(wb, {
      type: 'buffer',
      cellStyles: true
    })
    download(data, filename) 
  }

  _this.writeToLocalFile = (filename) => {
    XLSX.writeFile(wb, filename)
  }
}

const filterJSONObjects = (jsonObjects, keys) =>
  jsonObjects.map(obj =>
    Object.keys(keys).reduce((filteredObj, key) => {
      if (key in obj) filteredObj[key] = obj[key];
      return filteredObj;
    }, {})
  );

