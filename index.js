require('dotenv').config()
const creds = {private_key: process.env.PRIVATE_KEY, client_email: process.env.CLIENT_EMAIL}

const { GoogleSpreadsheet } = require('google-spreadsheet');
const express = require('express')
const app = express()
const port = process.env.PORT || 3000

const {annual_indices, template_indices} = require('./indices.json')
const test_entry = require('./entry.json');

const HYPERLINK = (sheet_id, text) => `=HYPERLINK("https://docs.google.com/spreadsheets/d/${sheet_id}", "${text}")`

// Initialize the sheet - doc ID is the long id in the sheets URL
const annual_budget = new GoogleSpreadsheet('1SmJMMxQ0hpVWBpl-dVw1hlJdNc11nNfQfAywEd6LG1c');
init()

async function init(){
    await annual_budget.useServiceAccountAuth(creds)
    await annual_budget.loadInfo()
}

app.get('/', async (req, res) => {
    const budgets = await parse_sheet('1nYsXA9Xl6K4yV0tHmaSUXyyvFrq9iU7TtG_tWvgoPHM')
    console.log(budgets)
    for(entry of budgets){
        await create_entry(entry)
    }

    res.send(`   
        Hello user! 
        This is a private service run by Sixth College Student Council's VP Finance. 
        Please direct questions to scsc.finance@gmail.com. Thanks! 
    `)
})

app.get('/:id', async (req, res) => {
    const budgets = await parse_sheet(req.params.id)
    console.log(budgets)
    for(entry of budgets){
        await create_entry(entry)
    }
    res.send(`   
        Hello user! 
        This is a private service run by Sixth College Student Council's VP Finance. 
        Please direct questions to scsc.finance@gmail.com. Thanks! 
    `)
})

app.listen(port, () => {
      console.log(`Example app listening at http://localhost:${port}`)
})

async function create_entry(entry){
    // load up all the cells in the annual budget
    console.time('cells loaded - time');
    funding_sheet = annual_budget.sheetsByTitle['Funding'];
    await funding_sheet.loadCells();
    console.timeEnd('cells loaded - time');

    // load the cell contents of the ID column
    const cells = []
    for(var i = 0; i < funding_sheet.rowCount; i++){
        cells.push(funding_sheet.getCell(i,0).value);
    }

    console.log(cells)

    // if the budget exists, update the row for it; otherwise, make a new row
    var idx = cells.indexOf(entry['id'])
    if(idx == -1)
        idx = cells.indexOf(null)

    console.log('index:', idx);

    // update all the cells 
    for(k in entry)
        if(k in annual_indices)
            funding_sheet.getCell(idx, annual_indices[k]).value = entry[k];

    // save changes to the sheet
    await funding_sheet.saveUpdatedCells();
}

async function parse_sheet(sheet_id){
    console.log(sheet_id)

    // Initialize the sheet - doc ID is the long id in the sheets URL
    const doc = new GoogleSpreadsheet(sheet_id);

    // Initialize Authentication - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth(creds)
    await doc.loadInfo();

    // load up all the cells 
    const sheet = doc.sheetsByTitle['PRE-Budget'];
    await sheet.loadCells();

    // make the budget object, set properties
    partial_budget = {}

    // set the name
    partial_budget['name'] = sheet.getCellByA1(template_indices.name).value

    // set the committee
    partial_budget['committee'] = sheet.getCellByA1(template_indices.committee).value

    // set the submitter's info
    partial_budget['submitter_name'] = sheet.getCellByA1(template_indices.submitter_name).value
    partial_budget['submitter_email'] = sheet.getCellByA1(template_indices.submitter_email).value

    // set the advisor's info
    partial_budget['advisor_name'] = sheet.getCellByA1(template_indices.advisor_name).value
    partial_budget['advisor_email'] = sheet.getCellByA1(template_indices.advisor_email).value

    // set the date info
    partial_budget['event_date'] = sheet.getCellByA1(template_indices.event_date).value
    partial_budget['date_submitted'] = new Date().toLocaleDateString('en-US', { timeZone: 'America/Los_Angeles' })

    // figure out what valid line items we're splitting between
    line_items = []
    for(a1 of template_indices.line_item_list){
        value = sheet.getCellByA1(a1).value
        if(value != null)
            line_items.push(value)
    }

    // make all the "real" budgets split between all the different line items (equal split)
    budgets = []

    const line_item_count = line_items.length
    const total_expense = sheet.getCellByA1(template_indices.expense).value 
    for(var i = 0; i < line_item_count; i++){
        copied_partial_budget = {...partial_budget};

        // set the total amount
        copied_partial_budget['expense'] = total_expense / line_item_count
        copied_partial_budget['line_item'] = line_items[i];
        copied_partial_budget['id'] = HYPERLINK(doc.spreadsheetId, `${doc.spreadsheetId}~~~~${i}`);

        budgets.push(copied_partial_budget);
    }

    return budgets
}

