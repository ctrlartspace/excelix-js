import XLSX from 'xlsx-js-style'
import download from 'downloadjs'


export default function Excelix() {
  let _this = this;
  let headersCount = 0
  let totalFields = null 

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet([{}])

  XLSX.utils.book_append_sheet(wb, ws, 'transactions')
  
  this.addRow = (row, { bold, inline=false, r=headersCount, c=0 }={}) => {
    console.log('add row: ', row)
    XLSX.utils.sheet_add_aoa(ws, [row], 
      { origin: 
        { 
          r: r,
          c: c
        } 
      })

    const cellRef = XLSX.utils.encode_cell({r:headersCount,c:c})
    console.log('style:', r, c)
    ws[cellRef].s = {font: {bold}}

    headersCount += 1
    //headersCount += inline ? 1 : 0 

  }


  
  this.addJson = (jsonData) => {
    if (!Array.isArray(jsonData)) {
      console.error('JSON data must be array')
      return
    }
    
    const filteredJsonData = filterJSONObjects(jsonData, 
      this.totalFields)  
    const titles = Object.keys(filteredJsonData[0])
    
    // add titles
    console.log(titles)
    console.log(this.totalFields)
    
    const titlesFormatted = []
    titles.forEach(e => {
       titlesFormatted.push(this.totalFields[e].text || e)
    })

    console.log(titlesFormatted)
    this.addRow(titlesFormatted, { bold: true }) 

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

    if (this.totalFields) {
      this.addRow(['ИТОГО'], { bold: true })
      const sums = [] 
      titles.forEach((c, i) => {
        const isTotal = this.totalFields[c].total

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

      })
      this.addRow(sums)
    }
  }

  this.addTotalFields = (totalFields) => {
    this.totalFields = totalFields
  }

  this.writeToFile = (filename) => {
    const data = XLSX.write(wb, {
      type: 'buffer',
      cellStyles: true
    })
    download(data, filename) 
  }

  this.writeToLocalFile = (filename) => {
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

