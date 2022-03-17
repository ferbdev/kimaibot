const puppeteer = require('puppeteer');
const Excel = require('exceljs');

var request = require('request');

require('dotenv').config();

console.log('Iniciando bot');

var userName = process.env.USER_NAME;
var password = process.env.USER_PASS;

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function getFormattedDate() {
    var date = new Date();
    var str = date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds() + "." + date.getMilliseconds() + "- ";

    return str;
}

function DoLog(message) {
    console.log(getFormattedDate() + message);
}

function getRandomNumberBetween(min, max) {
    return Math.floor(Math.random() * (max - min + 1) + min);
}

function pad2(n) { return n < 10 ? '0' + n : n }

var date = new Date();

function getStringDate(){
    var stringDate = date.getFullYear().toString() + pad2(date.getMonth() + 1) + pad2( date.getDate()) + pad2( date.getHours() ) + pad2( date.getMinutes() ) + pad2( date.getSeconds() ) + pad2( date.getMilliseconds() );
    return stringDate;
}

async function StartBot(browser) {

    const pageBot = await browser.newPage();

    await pageBot.setUserAgent(
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"
    );
	
	await pageBot.setExtraHTTPHeaders({
	 'Accept-Language': 'en'
	});

    await pageBot.goto('https://platform.senior.com.br/login/?redirectTo=https%3A%2F%2Fplatform.senior.com.br%2Fplataforma%2F&tenant=knapp.com', {
        waitUntil: 'networkidle0',
    });

    console.log('Fazendo login...');

    var user = await pageBot.evaluate((userName) => {
        return document.getElementById('username-input-field').value = userName;
    }, userName);

    await sleep(100);

    var pass = await pageBot.evaluate((password) => {
        return document.getElementById('password-input-field').value = password;
    }, password);

    await sleep(100);

    var login = await pageBot.evaluate(() => {
        return document.getElementById('loginbtn').click();
    });

    while(pageBot.url().includes("https://platform.senior.com.br/login/")){
        console.log('Waiting login...');
        await sleep(500);
    }

    console.log('Getting appointments...');

    var token;
    let after = await pageBot.cookies();

    var stringJson;
    var userFullName;

    after.forEach(cookie => {

        if(cookie.name == 'com.senior.pau.token'){
            stringJson = JSON.parse(cookie.value);

            console.log(stringJson);
        }

        if(cookie.name == 'com.senior.pau.userdata'){
            var userData = cookie.value;

            var index0 = userData.indexOf("fullname") + 17;

            if(index0 <= 17){
                return;
            }  

            var index1 = userData.indexOf("description") - 9;

            userFullName = userData.substring(index0, index1).replaceAll('+', ' ');

            console.log("name", userFullName);
        }
    })

    var cookie = await pageBot.evaluate((stringJson) => {
        return JSON.parse(stringJson).access_token;
    }, stringJson);

    console.log(cookie);

    token = 'bearer ' + cookie;

    console.log('Cookie found', token);

    var options = {
    'method': 'POST',
    'url': 'https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/hcm/pontomobile/queries/clockingEventByActiveUserQuery',
    'headers': {
        'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'Authorization': token,
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36',
        'sec-ch-ua-platform': '"Windows"'
    },
    body: JSON.stringify({
        "filter": {
        "activePlatformUser": true,
        "pageInfo": {
            "page": 0,
            "pageSize": "50"
        },
        "nameSearch": userFullName,
        "sort": {
            "field": null,
            "order": "ASC"
        }
        }
    })

    };
    request(options, function (error, response) {
        if (error) throw new Error(error);

        var appointmentsObj = JSON.parse(response.body);

        var startRow = 5;
        var startColumn = 1;

        const fileName = 'apontamento_temp.xlsx';

        const wb = new Excel.Workbook();

        const ws = wb.addWorksheet('My Sheet');

        var grouped = groupBy(appointmentsObj.result, app => app.dateEvent);

        let result = Array.from(grouped).sort(sortByProperty(0));

        //console.log(result);

        result.forEach(element => {

            ws.getCell(`B${startRow}`).value = FormataStringData(element[1][0].dateEvent);

            for (var i = element[1].length; i > 0; i--) {

                ws.getCell(`${String.fromCharCode(97 + (element[1].length - i) + 2).toUpperCase()}${startRow}`).value = element[1][i-1].timeEvent;
            }

            startRow++;
        });

        wb.xlsx
        .writeFile(fileName)
        .then(() => {
            console.log('file created');
        })
        .catch(err => {
            console.log(err.message);
        });
    });


    pageBot.close();

}

function groupBy(list, keyGetter) {
    const map = new Map();
    list.forEach((item) => {
         const key = keyGetter(item);
         const collection = map.get(key);
         if (!collection) {
             map.set(key, [item]);
         } else {
             collection.push(item);
         }
    });
    return map;
}

function sortByProperty(property){  
    return function(a,b){  
       if(a[property] > b[property])  
          return 1;  
       else if(a[property] < b[property])  
          return -1;  
   
       return 0;  
    }
}

 function FormataStringData(data) {
    var ano  = data.split("-")[0];
    var mes  = data.split("-")[1];
    var dia  = data.split("-")[2];
  
    return ("0"+dia).slice(-2) + '/' + ("0"+mes).slice(-2) + '/' + ano;
    // Utilizo o .slice(-2) para garantir o formato com 2 digitos.
  }

async function InitBot(){
    const browser = await puppeteer.launch({
        headless: true,
    });  

    StartBot(browser);
}

InitBot();