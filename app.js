/**
 * Created by apetrov on 02.06.2017.
 */
/**
 * Created by PhpStorm.
 * Company: Appalachian Ltd.
 * Developer: SETTER
 * Suite: appalachi.ru
 * Email: info@appalachi.ru
 * Date: 01.06.2017
 * Time: 22:55
 */

function SKD() {
    let name = this.name;
    let dateCreate = this.dateCreate;
    let startPeriod = this.startPeriod;
    let endPeriod = this.endPeriod;
    let headerOneRow = this.headerOneRow;
    let headerTwoRow = this.headerTwoRow;
    let data = [];


}
SKD.prototype.checkedTypeFile = function () {

};

// const request = require('request');
const XLSX = require('xlsx');
const XlsxPopulate = require('xlsx-populate');
const http = require('http');
const fs = require('fs');
const Watcher = require('listener-dir');
const execFile = require('child_process').execFile;


/**
 * Путь до папки с файлами .xlsx из SKD
 */
const sourceReportSkd = 'F:/!_NODE_PROJECTS/KADR/skd-report/';
const targetReportSkd = './skd/xlsx/';


/**
 * Поток чтения файла
 */
// var myReadStream = fs.createReadStream(sourceReportSkd + fileName);

/**
 * Поток записи файла
 */
// var myWriteStream = fs.createWriteStream(targetReportSkd + fileName);


/**
 *  После определения класса Watcher можно воспользоваться им,
 *  создав объект Watcher
 *  Первый аргумен папка для прослушивания
 *  Второй аргумент папка назначения файла. Куда он будет перемещён.
 * @type {Watcher}
 */
var watcher = new Watcher(sourceReportSkd, targetReportSkd);


/**
 * В только что созданном объекте Watcher можно использовать метод on,
 * унаследованный от класса генератора событий, чтобы создать
 * логику обработки каждого файла,
 */
watcher.on('process', function process(file) {
    var pt = this.watchDir;
    var watchFile = this.watchDir + '/' + file;
    var processedFile = this.processedDir + '/' + file;


    execFile('file', ['-b', '--mime-type', watchFile], function (error, stdout, stderr) {
        if (stdout.trim() === 'application/vnd.ms-office' || stdout.trim() === 'application/zip') {
            if (file.slice(-3) == 'xls') {
                var nameFile = file.slice(0, -3);
                console.log(nameFile);
                var workbook2 = XLSX.readFile(watchFile);

                var first_sheet_name = workbook2.SheetNames[0];

                /* Get worksheet */
                var worksheet = workbook2.Sheets[first_sheet_name];
                console.log(XLSX.utils.sheet_to_json(worksheet, {header: ["A", "B", "C", "D", "E", "F"]}));
                // console.log(XLSX.utils.sheet_to_json(worksheet, {header:["A","E","I","O","U","6","9"]}));
                // console.log(XLSX.utils.sheet_to_json(worksheet, {header:1}));
                // console.log(XLSX.utils.sheet_to_json(worksheet, {header:"A"}));

                XLSX.writeFile(workbook2, pt + nameFile + 'xlsx');
                fs.unlink(pt + nameFile + file.slice(-3), (err)=> {
                    "use strict";
                    if (err) throw err;
                })
            }
            if (file.slice(-4) == 'xlsx') {
                fs.rename(watchFile, processedFile, (err)=> {
                    if (err) console.error("Server Error" + err);
                    XlsxPopulate.fromFileAsync(processedFile).then(function (workbook) {
                            "use strict";
                            /**
                             * Матрица вся книга.
                             * @type {Range|undefined}
                             */
                            const matrix = workbook.sheet(0).usedRange();
                            // console.log(matrix);
                        },
                        function (reject) {
                            console.log('Error reject: ' + reject);
                        });
                });
            }
        }
    });
});


/**
 * Теперь, после создания всего необходимого кода, инициировать мониторинг
 * папки можно с помощью такой команды:
 */
watcher.start();
//
// myReadStream.on('data', function (chunk) {
//     console.log('Получен новый объем данных: ');
//     myWriteStream.write(chunk);
// });
//
// myReadStream.on('finish', function (chunk) {
//     "use strict";
//     console.log("finish");
// });
//
// myReadStream.on('error', function (err) {
//     // res.statusCode = 500;
//     // res.end("Server Error");
//     console.error("Server Error" +err);
// });
//
// myReadStream.on('open', function () {
//     console.log("open");
// }).on('close', function () {
//         console.log("close");
//         XlsxPopulate.fromFileAsync(targetReportSkd + fileName).then(function (workbook) {
//                 "use strict";
//                 /**
//                  * Матрица вся книга.
//                  * @type {Range|undefined}
//                  */
//                 const matrix = workbook.sheet(0).usedRange();
//                 console.log(matrix);
//             },
//             function (reject) {
//                 console.log('Prommissss 45 ' + error);
//             });
//     });
//


//********************************************


//fs.open('./' + fileName, 'w', (err, fd) => {
//    if (err) console.error(err);
//
//res = fd;
//});
//const defaults = {
//    flags: 'r',
//    encoding: null,
//    fd: res,
//    mode: 0o666,
//    autoClose: true
//};
//var http = require('http');
//var fs = require('fs');


//var stream = new fs.ReadStream('Z:\\Landata 2017-05-30.xls', {encoding: 'utf-8'});

//var dest = stream.pipe(fs.createWriteStream('./' + fileName));


//var promise = new Promise((resolve, reject) => {
//    dest.on('finish', resolve);
//    source.on('error', reject);
//    dest.on('error', reject);
//}).catch(e => new Promise((_, reject) => {
//    dest.end(() => {
//        fs.unlink('doodle.png', () => reject(e));
//    });
//}));


//sendFile(stream, dest);

// var stream = new fs.ReadStream('Z:\\' + fileName);
// stream.on('readable', function () {
//     var data = stream.read();
//     if (data != null) console.log(data.length / 1000 + 'Kb');
//     //console.log(data);
// });
//
//
// stream.on('end', function () {
//     console.log("THE END");
// });

/**
 * Чтобы Node не упал ставим,обрабочик на ошибку.
 */
// stream.on('error', function (err) {
//     if (err.code == 'ENOENT') {
//         console.log("Файл не найден");
//     } else {
//         console.error(err);
//     }
// });
//
// function sendFile(file, res) {
//     file.pipe(res);
//
//     file.on('error', function (err) {
//         res.statusCode = 500;
//         res.end("Server Error");
//         console.error(err);
//     });
//
//     file
//         .on('open', function () {
//             console.log("open");
//         })
//         .on('close', function () {
//             console.log("close");
//         });
//
//     res.on('close', function () {
//         file.destroy();
//     });
// }


//XlsxPopulate.fromFileAsync(data).then(function (workbook) {
//        "use strict";
//        /**
//         * Матрица вся книга.
//         * @type {Range|undefined}
//         */
//        const matrix = workbook.sheet(0).usedRange();
//        sails.log(matrix);
//    },
//    function (reject) {
//        sails.log('Prommissss 45 ' + error);
//    });


// fs.readdir('Z:\\', {encoding: 'utf-8'}, function (err, data) {
//     if (err) {
//         console.log(err);
//     } else {
//         //console.log(data[4]);
//
//         console.log(data);
//     }
// });


//var smb2Client = new SMB2({
//    share:'\\\\portal\\DavWWWRoot\\it\\skd\\'
//    , domain:'landata'
//    , username:'apetrov@portal'
//    , password:'Tel-4211817'
//});
////smb2Client.exists('skd\\', function (err, exists) {
////    if (err) throw err;
////    console.log(exists ? "it's there" : "it's not there!");
////});
//smb2Client.readdir('Windows\\System32', function(err, files){
//    if(err) throw err;
//    console.log(files);
//});

//smb2Client.readdir('skd', function(err, files){
//    if(err) throw err;
//    console.log(files);
//});

//
// var watch = 'Z:\\';
//var watch = 'watch';
// После определения класса Watcher можно воспользоваться им, создав объект Watcher
// var watcher = new Watcher(watch, './done');

/**
 * В только что созданном объекте Watcher можно использовать метод on,
 * унаследованный от класса генератора событий, чтобы создать
 * логику обработки каждого файла,
 */
// watcher.on('process', function process(file) {
//     var watchFile = this.watchDir + '/' + file;
//     var processedFile = this.processedDir + '/' + file.toLowerCase();
//     fs.rename(watchFile, processedFile, (err) => {
//         if (err) throw err;
//     })
//     ;
// });

/**
 * Теперь, после создания всего необходимого кода, инициировать мониторинг
 * папки можно с помощью такой команды:
 */
// watcher.start();


// portal/DavWWWRoot/it/skd
//fs.watch('Z:\\1.txt', (eventType, filename) => {
//    console.log(`event type is: ${eventType}`);
//    if (filename) {
//        console.log(`filename provided: ${filename}`);
//    } else {
//        console.log('filename not provided');
//    }
//});
//fs.watch('\\\\portal\\DavWWWRoot\\it\\skd\\'+data[4], (eventType, filename) => {
//    console.log(`event type is: ${eventType}`);
//    if (filename) {
//        console.log(`filename provided: ${filename}`);
//    } else {
//        console.log('filename not provided');
//    }
//});