const { Router, response } = require("express");
const Excel = require('exceljs')


//const assert = require("chai").assert;

module.exports = function(app, fs)
{
//var initiFrameUrl = "https://pack.kt.com/vcoloring/m/index2.asp"
//var initiFrameUrl = "https://zone.membership.kt.com/event/2021NewGreen/m/index.asp"
//var initiFrameUrl = "https://zone.membership.kt.com/event/2021Petsbe/m/"





//2021.04.28 excel 모듈 테스트 start ***********************************************************/
app.get('/readExcel3', function(req, res){

  var workbook = new Excel.Workbook(); 
  workbook.xlsx.readFile('ykk.xlsx')
      .then(function() {
          var worksheet = workbook.getWorksheet('Debtors');
          worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
          });
      });

})

app.get('/readExcel2', function(req, res){
var workbook = new Excel.Workbook();
workbook.creator ="Naveen"; 
workbook.modified ="Kumar";
workbook.xlsx.readFile("ykk.xlsx").then(function(){
    var workSheet =  workbook.getWorksheet("Debtors"); 

    workSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {

        currRow = workSheet.getRow(rowNumber); 
         console.log("User Name :" + currRow.getCell(1).value +", Password :" +currRow.getCell(2).value);
         console.log("User Name :" + row.values[1] +", Password :" +  row.values[2] ); 

      //   assert.equal(currRow.getCell(2).type, Excel.ValueType.Number); 
       //  console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
      });
      
})
})


app.get('/readExcel', function(req, res){
  var workbook = new Excel.Workbook();
  workbook.creator ="Naveen"; 
  workbook.modified ="Kumar";
  workbook.xlsx.readFile("Debtors.xlsx").then(function(){
      var workSheet =  workbook.getWorksheet("Debtors"); 
  var vfData = {}
      workSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  
          currRow = workSheet.getRow(rowNumber); 

           console.log("mainjs User Name :" + currRow.getCell(1).value +", Password :" +currRow.getCell(2).value);
           console.log("mainjsUser Name :" + row.values[1] +", Password :" +  row.values[2] ); 
  
        //   assert.equal(currRow.getCell(2).type, Excel.ValueType.Number); 
         //  console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
        });
        
  })

  res.render('readExcel', {
    data : 'a b ykk',
     data2 : '가 나 다'
  })
})


app.get('/',function(req,res){
  res.render('index', {
      title: "KT Event",
      length: 5,
//             iframeUrl : "https://zone.membership.kt.com/event/2021Lotteworld"
      iframeUrl : initiFrameUrl,
      iframeHeight : initiframeHeight
  })
});









app.get('/createExcel', function(req, res){
    'use strict'

    const data = [{
      firstName: 'John',
      lastName: 'Bailey',
      purchasePrice: 1000,
      paymentsMade: 100
    }, {
      firstName: 'Leonard',
      lastName: 'Clark',
      purchasePrice: 1000,
      paymentsMade: 150
    }, {
      firstName: 'Phil',
      lastName: 'Knox',
      purchasePrice: 1000,
      paymentsMade: 200
    }, {
      firstName: 'Sonia',
      lastName: 'Glover',
      purchasePrice: 1000,
      paymentsMade: 250
    }, {
      firstName: 'Adam',
      lastName: 'Mackay',
      purchasePrice: 1000,
      paymentsMade: 350
    }, {
      firstName: 'Lisa',
      lastName: 'Ogden',
      purchasePrice: 1000,
      paymentsMade: 400
    }, {
      firstName: 'Elizabeth',
      lastName: 'Murray',
      purchasePrice: 1000,
      paymentsMade: 500
    }, {
      firstName: 'Caroline',
      lastName: 'Jackson',
      purchasePrice: 1000,
      paymentsMade: 350
    }, {
      firstName: 'Kylie',
      lastName: 'James',
      purchasePrice: 1000,
      paymentsMade: 900
    }, {
      firstName: 'Harry',
      lastName: 'Peake',
      purchasePrice: 1000,
      paymentsMade: 1000
    }]
    
   // const Excel = require('exceljs')
    
    // need to create a workbook object. Almost everything in ExcelJS is based off of the workbook object.
    let workbook = new Excel.Workbook()
    
    let worksheet = workbook.addWorksheet('Debtors')
    
    worksheet.columns = [
      {header: 'First Name', key: 'firstName'},
      {header: 'Last Name', key: 'lastName'},
      {header: 'Purchase Price', key: 'purchasePrice'},
      {header: 'Payments Made', key: 'paymentsMade'},
      {header: 'Amount Remaining', key: 'amountRemaining'},
      {header: '% Remaining', key: 'percentRemaining'}
    ]
    
    // force the columns to be at least as long as their header row.
    // Have to take this approach because ExcelJS doesn't have an autofit property.
    worksheet.columns.forEach(column => {
      column.width = column.header.length < 12 ? 12 : column.header.length
    })
    
    // Make the header bold.
    // Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
    worksheet.getRow(1).font = {bold: true}
    
    // Dump all the data into Excel
    data.forEach((e, index) => {
      // row 1 is the header.
      const rowIndex = index + 2
    
      // By using destructuring we can easily dump all of the data into the row without doing much
      // We can add formulas pretty easily by providing the formula property.
      worksheet.addRow({
        ...e,
        amountRemaining: {
          formula: `=C${rowIndex}-D${rowIndex}`
        },
        percentRemaining: {
          formula: `=E${rowIndex}/C${rowIndex}`
        }
      })
    })
    
    const totalNumberOfRows = worksheet.rowCount
    
    // Add the total Rows
    worksheet.addRow([
      '',
      'Total',
      {
        formula: `=sum(C2:C${totalNumberOfRows})`
      },
      {
        formula: `=sum(D2:D${totalNumberOfRows})`
      },
      {
        formula: `=sum(E2:E${totalNumberOfRows})`
      },
      {
        formula: `=E${totalNumberOfRows + 1}/C${totalNumberOfRows + 1}`
      }
    ])
    
    // Set the way columns C - F are formatted
    const figureColumns = [3, 4, 5, 6]
    figureColumns.forEach((i) => {
      worksheet.getColumn(i).numFmt = '$0.00'
      worksheet.getColumn(i).alignment = {horizontal: 'center'}
    })
    
    // Column F needs to be formatted as a percentage.
    worksheet.getColumn(6).numFmt = '0.00%'
    
    // loop through all of the rows and set the outline style.
    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      worksheet.getCell(`A${rowNumber}`).border = {
        top: {style: 'thin'},
        left: {style: 'thin'},
        bottom: {style: 'thin'},
        right: {style: 'none'}
      }
    
      const insideColumns = ['B', 'C', 'D', 'E']
    
      insideColumns.forEach((v) => {
        worksheet.getCell(`${v}${rowNumber}`).border = {
          top: {style: 'thin'},
          bottom: {style: 'thin'},
          left: {style: 'none'},
          right: {style: 'none'}
        }
      })
    
      worksheet.getCell(`F${rowNumber}`).border = {
        top: {style: 'thin'},
        left: {style: 'none'},
        bottom: {style: 'thin'},
        right: {style: 'thin'}
      }
    })
    
    // The last A cell needs to have some of it's borders removed.
    worksheet.getCell(`A${worksheet.rowCount}`).border = {
      top: {style: 'thin'},
      left: {style: 'none'},
      bottom: {style: 'none'},
      right: {style: 'thin'}
    }
    
    const totalCell = worksheet.getCell(`B${worksheet.rowCount}`)
    totalCell.font = {bold: true}
    totalCell.alignment = {horizontal: 'center'}
    
    // Create a freeze pane, which means we'll always see the header as we scroll around.
    worksheet.views = [
      { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' }
    ]
    
    // Keep in mind that reading and writing is promise based.
    workbook.xlsx.writeFile('ykk.xlsx')
    
    res.send('create excel')

})

//2021.04.28 excel 모듈 테스트 end ***********************************************************


var initiFrameUrl = "https://kt-promotion.com/passbykt/210401_02/?route=ktcom"

var initiframeHeight= "20000"

     app.get('/',function(req,res){
         res.render('index', {
             title: "KT Event",
             length: 5,
//             iframeUrl : "https://zone.membership.kt.com/event/2021Lotteworld"
             iframeUrl : initiFrameUrl,
             iframeHeight : initiframeHeight
         })
     });

 
     app.post('/iframe', function(req, res){
        var iframeUrl = req.body.iframeUrl;
        var iframeHeight = req.body.iframeHeight;
        
        console.log(iframeUrl);
        res.render('index', {
            title: "KT Event",
            length: 5,
            iframeUrl : iframeUrl,
            iframeHeight : iframeHeight
        })
     });

     app.get('/iframe', function(req, res){
        var iframeUrl = req.body.iframeUrl;
        console.log(iframeUrl);
            res.render('index', {
            title: "KT Event",
            length: 5,
            iframeUrl : initiFrameUrl,
            iframeHeight : initiframeHeight
        })
     });

    

     app.get('/help', function(req, res){
        var iframeUrl = req.body.iframeUrl;
        console.log(iframeUrl);
            res.render('help', {
            
        })
     });


    app.get('/list', function (req, res) {
       fs.readFile( __dirname + "/../data/" + "user.json", 'utf8', function (err, data) {
           console.log( data );
           res.end( data );
       });
    })

app.get('/getUser/:username', function(req, res){
    fs.readFile( __dirname + "/../data/user.json", 'utf8', function (err, data) {
         var users = JSON.parse(data);
         res.json(users[req.params.username]);
    });
 });

//  app.post('/iframe', function(req, res){
//     var iframeUrl = req.body.iframeUrl;
//     console.log(iframeUrl);
//     res.render('index', {
//         title: "MY HOMEPAGE",
//         length: 5,
//         iframeUrl : iframeUrl
//     })
    
//  });


 /*
 app.post('/form_receiver', function(req, res){
    var title = req.body.title;
    var description = req.body.description;
    res.send(title+','+description);
  });
*/

app.post('/addUser/:username', function(req, res){
    var result = {  };
    var username = req.params.username;

    // CHECK REQ VALIDITY
    if(!req.body["password"] || !req.body["name"]){
        result["success"] = 0;
        result["error"] = "invalid request";
        res.json(result);
        return;
    }

    // LOAD DATA & CHECK DUPLICATION
    fs.readFile( __dirname + "/../data/user.json", 'utf8',  function(err, data){
        var users = JSON.parse(data);
        if(users[username]){
            // DUPLICATION FOUND
            result["success"] = 0;
            result["error"] = "duplicate";
            res.json(result);
            return;
        }

        // ADD TO DATA
        users[username] = req.body;

        // SAVE DATA
        fs.writeFile(__dirname + "/../data/user.json",
                     JSON.stringify(users, null, '\t'), "utf8", function(err, data){
            result = {"success": 1};
            res.json(result);
        })
    })
});


app.put('/updateUser/:username', function(req, res){
  var result = {};
  var username = req.params.username;
  //check req validity
  if (!req.body["password"] || !req.body["name"]) {
    result["success"] = 0;
    result["error"] = "invalid request" ;
    res.json(result);
    return;
  }

  fs.readFile(__dirname + "/../data/user.json", "utf8", function(err, data){
      var users = JSON.parse(data);
      //ADD / MODIFY DATA
        users[username] = req.body;

        // SAVE DATA
        fs.writeFile(__dirname + "/../data/user.json", JSON.stringify(users, null, '\t'), "utf8", function(err, data){
            result = {"success" : 1};
            res.json(result);
  })        

  })
});

app.delete('/deleteUser/:username', function(req,res){
    var result = { }
    fs.readFile(__dirname + "/../data/user.json", "utf8", function(err,data){
        var users = JSON.parse(data);
        //IF NOT FOUND
        if(!users[req.params.username]){
            result["success"] = 0;
            result["error"] = "not found";
            res.json(result);
            return;
        }
 
    delete users[req.params.username];
    fs.writeFile(__dirname + "/../data/user.json", JSON.stringify(users, null, '\t'), "utf8", function(err, data){
        result["success"] = 1;
        res.json(result);
        return;
    })
  })
})


}