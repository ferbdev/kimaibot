var xlsx = require('node-xlsx');

var request = require('request');


const IdProjetoMultiB = 1022;
const IdProjetoAlmoco = 1041;

const ActivityIDMultiB = 39;
const ActivityIDAlmoco = 10;

let kimayKey;

let user = '';
let password = '';

const ReadExcel = async () => {
  var obj = xlsx.parse(__dirname + '/apontamento.xlsx'); // parses a file

  var objApontamentos = obj[0].data;

  user = objApontamentos[1][1];
  password = objApontamentos[2][1];

  console.log('Fazendo login...');
  console.log(user, password);

  let login = KimaiLogin();
  let response_login = await login;

  console.log(response_login);

  var apontamentos = [];
  for(var i=5; i < objApontamentos.length; i++){
    if(objApontamentos[i].length > 1) 
      apontamentos.push(objApontamentos[i]);
  }

  //console.log(apontamentos);
  
  for(var i=0; i < apontamentos.length; i++){
    var day = apontamentos[i][0].replace('/', '.').replace('/', '.');
    
    for(var j=1; j < apontamentos[i].length - 1; j++){
      
      var entrada = apontamentos[i][j];
      var saida = apontamentos[i][j + 1];

      var isAlmoco = (j == 2);

      var total = j / 2;
      var ignore = total % 2 == 0 ? true : false;

      if(ignore && !isAlmoco){
        continue;
      }

      idProj = isAlmoco ? IdProjetoAlmoco : IdProjetoMultiB;
      idActv = isAlmoco ? ActivityIDAlmoco : ActivityIDMultiB;

      let appointment = KimaiAppoint(idProj, idActv, entrada, saida, day);
      let response_body = await appointment;

      console.log(response_body);

      console.log(`Apontou proj ${idProj} act ${idActv} || ${day} das ${entrada} as ${saida}`)
    }
    
  }
}

const KimaiLogin = async () => {
  return new Promise(async (resolve, reject) => {
    var options = {
      'method': 'POST',
      'url': 'https://timetracking.knapp.com.br/index.php?a=checklogin',
      'headers': {
        'Cookie': 'kimai_key=iH7M0VTyLSzIWeuherk1VtfGGGPFM2; kimai_user=lucas.bueno'
      },
      formData: {
        'name': user,
        'password': password
        }
    };
    await request(options, async function (error, response) {
      if (error) console.log('exception:' + error);

      kimayKey = response.rawHeaders.filter(name => name.includes('kimai_key='));
      resolve('Logado!');
    });
  })
}

const KimaiAppoint = (idproj, activity, startTime, endTime, day) => {

  return new Promise(async (resolve, reject) => {
    var options = {
      'method': 'POST',
      'url': 'https://timetracking.knapp.com.br/extensions/ki_timesheets/processor.php',
      'headers': {
        'sec-ch-ua': '"Chromium";v="92", " Not A;Brand";v="99", "Google Chrome";v="92"',
        'Accept': '*/*',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': `${kimayKey}; kimai_user=lucas.bueno`
      },
      form: {
        'activityID': activity,
        'approved': '',
        'axAction': 'add_edit_timeSheetEntry',
        'billable': '0',
        'budget': '',
        'comment': '',
        'commentType': '0',
        'description': '',
        'duration': '00:00:00',
        'end_day': day,//'21.08.2021',
        'end_time': endTime, //12:00:00
        'filter': '',
        'filter': '',
        'id': '',
        'location': '',
        'projectID': idproj,
        'start_day': day,//'21.08.2021',
        'start_time': startTime, //08:00:00
        'statusID': '1',
        'trackingNumber': '',
        'userID[]': '932335395' 
      }
    };
    
    await request(options, function (error, response) {
      if (error) throw new Error(error);
      resolve(response.body);
    });
  });
}

ReadExcel();

/*setInterval(async ()=>{
    
    //get data from ponto rep
    //await processTrade();

}, 1000);*/

