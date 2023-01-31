import XLSX from 'xlsx-js-style'
import axios from 'axios'
import download from 'downloadjs'
import fs from 'fs'


export default function Excelix() {
  let _this = this;
  let headersCount = 0
  let totalFields = null 

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet([{}])

  XLSX.utils.book_append_sheet(wb, ws, 'transactions')
  
  this.addHeader = (header) => {
    // const row = [{ v: header, t: 's', s: { font: { bold: true }}}]
    // XLSX.utils.aoa_to_sheet([row])
    XLSX.utils.sheet_add_aoa(ws, [[header]], 
      { origin: 
        { 
          r: headersCount,
          c: 0
        } 
      })

    const cellRef = XLSX.utils.encode_cell({r:headersCount,c:0})
    ws[cellRef].s = {font: {bold: true}}

    headersCount += 1

  }


  
  this.addJson = (jsonData) => {
    const totalKeys = Object.keys(this.totalFields)
    const jsonColumns = Object.keys(jsonData[0])
    XLSX.utils.sheet_add_aoa(ws, [jsonColumns],
      {
        origin:
        {
          r: headersCount,
          c: 0
        }
      })

    headersCount += 1

    XLSX.utils.sheet_add_json(ws, jsonData,
      { origin:
        {
          r: headersCount,
          c: 0
        },
        skipHeader: true

      })
    headersCount += Object.keys(jsonData).length

    if (this.totalFields) {


      // add Total title
      XLSX.utils.sheet_add_aoa(ws, [['Итого']],
        {
          origin:
          {
            r: headersCount,
            c: 0
          }
        })

      headersCount += 1

      // add Total values for each column 
      // in totalFields
      jsonColumns.forEach((c, i) => {
        console.log(c)
        // проверяем, если есть поле 
        // если есть, смотрим, нужно ли суммировать поле
        if (totalKeys.includes(c) && this.totalFields[c].total) {
          let sum = 0
          jsonData.forEach(d => {
            if (!isNaN(d[c])) {
              sum += d[c]
            }
          })

          XLSX.utils.sheet_add_aoa(ws, [[sum]], 
            { origin: 
              {
                r: headersCount,
                c: i 

              }
            })

          const cellRef = XLSX.utils.encode_cell({r:headersCount,c:i})
          ws[cellRef].s = {font: {bold: true}}

        }
      })
      
      headersCount += 1
    }

  }

  this.addFooter = (footer) => {
    XLSX.utils.sheet_add_aoa(ws, [[footer]], 
      { origin: 
        { 
          r: headersCount,
          c: 0
        } 
      })
    
    const cellRef = XLSX.utils.encode_cell({r:headersCount,c:0})
    ws[cellRef].s = {font: {bold: true}}

    headersCount += 1
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
    //try {
    //  fs.writeFileSync(filename, data)
    //} catch (err) {
    //  console.error(err)
    //}
  }



}

async function getData() {
  const url = 'https://jsonplaceholder.typicode.com/todos'
  const res = await axios.get(url)
  return res.data
}


async function main() {
  const data = await getData()
  
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet([{}])
  XLSX.utils.sheet_add_json(ws, [{}], { header: ['1', '2', '3'], origin: 'A1' })
  XLSX.utils.sheet_add_json(ws, data, { origin: 'A2' })
  XLSX.utils.book_append_sheet(wb, ws, 'transactions')

  XLSX.writeFile(wb, 'trans.xlsx')

}


function main2() {
  const ex = new Excelix() 
  ex.addHeader('test')
  ex.addHeader('test2')
  //const jsonData = await getData()
  const totalFields = {}
  totalFields.id = { text: 'ID', total: true }
  totalFields.title = { text: 'Название' }
  ex.addTotalFields(totalFields)
  //ex.addJson(jsonData)
  ex.addFooter('footer')
  ex.addFooter('footer')
  ex.writeToFile('test.xlsx')
}


main2()
