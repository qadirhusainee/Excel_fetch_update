const Excel = require('exceljs');
const axios = require('axios');
const Promise = require("promise")


let workbook = new Excel.Workbook();
let filename = 'vimeo_id.xlsx';

workbook.xlsx.readFile(filename).then(function () {

  let worksheet = workbook.getWorksheet(1);
  let colA = worksheet.getColumn('A');
  let colB = worksheet.getColumn('B');
  let vimeoIDs = colA.values.splice(2);
  Promise.all(vimeoIDs.map(id => axios.get(`https://player.vimeo.com/video/${id}/config`))).then((respList) => {

    let thumbnails = respList.map((resp) => resp.data.video.thumbs["640"]);
    colB.values = [...colB.values.slice(1, 2), ...thumbnails];
    
    // Save the content in a new file (formulas re-calculated)
    workbook.xlsx.writeFile(__dirname + '/thumbnails.xlsx')
      .then(function () {
        console.log('Successfully completed');
      }).catch((error) => {
        console.log("Some error occurred")
      })
  }).catch((error) => {
    console.log("Some error occurred")
  })

});