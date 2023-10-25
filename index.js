import pgPromise from 'pg-promise'
import XLSX from 'xlsx'
import express from 'express'
import dotenv from 'dotenv'

dotenv.config()

const app = express()

// Database connection parameters
const dbConfig = {
  user: process.env.DB_USER,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  host: process.env.DB_HOST,
  port: process.env.DB_PORT,
}

const pgp = pgPromise()
const db = pgp(dbConfig)

// SQL query
const query = 'SELECT * FROM "time_trackings"'

const saveDataToXslx = (data) => {
  // Create a new workbook
  const workbook = XLSX.utils.book_new()
  const wsName = 'Sheet1'

  // Create a worksheet
  const wsData = [Object.keys(data[0]), ...data.map((row) => Object.values(row))]
  const ws = XLSX.utils.aoa_to_sheet(wsData)

  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, ws, wsName)

  // Create a binary buffer from the XLSX data
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })

  return buffer
}

app.get('/download-xlsx', async (req, res) => {
  try {
    const data = await db.any(query)

    // Set response headers for download
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.set('Content-Disposition', 'attachment; filename=report.xlsx')

    const xlsxFile = saveDataToXslx(data)
    res.send(xlsxFile)
  } catch (error) {
    console.error('ERROR:', error)
    res.status(500).send('Internal Server Error')
  }
})

const port = process.env.SERVER_PORT || 3333
app.listen(port, () => console.log(`listening on: http://localhost:${port}`))
