const express = require('express');
const multer = require('multer');
const excelToJson = require('convert-excel-to-json');
const mongodb = require('mongodb').MongoClient;

const upload = multer({dest: 'uploads/'});
const app = express();

const mongo_db = 'user_db';
const mongo_coll = 'user_coll';
const mongo_url = `mongodb://localhost/${mongo_db}`;

renameKeys = (keysMap, obj) => Object.keys(obj).reduce((acc, key) => {
    return {...acc, ...{[keysMap[key] || key]: obj[key]}}
}, {});

parseExcelResult = (result) => {
    result = result["Sheet1"];

    let header = null;
    let end = false;

    let final_result = [];

    result.forEach((value, idx) => {
        if (!end) {
            Object.keys(value).forEach(key => {
                if (header === null) {
                    if (value[key] === 'Systemname') {
                        header = value;
                    }
                } else {
                    if (value[key] === 'Virtualisieren FFM') {
                        end = true;
                    }
                }
            });

            if (header && !end) {
                final_result.push(renameKeys(header, value));
            }
        }
    });
    final_result.shift();
    return final_result;
};

mongoInsertResults = (data) => {
    mongodb.connect(mongo_url, function (err, client) {
        if (err) {
            throw err;
        }

        const db = client.db(mongo_db);

        db.collection(mongo_coll).insertMany(data, function (err, res) {
            if (err) {
                throw err;
            }
            console.log("Number of documents inserted: " + res.insertedCount);
        });
    });
};

mongoRespondRecords = (res) => {
    mongodb.connect(mongo_url, function (err, client) {
        if (err) {
            throw err;
        }

        let db = client.db(mongo_db);

        db.collection(mongo_coll).find().toArray(function (err, result) {
            if (err) {
                throw err;
            }

            res.json(result);
        });
    });
};

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.get('/mongo', (req, res) => {
    mongoRespondRecords(res);
});

app.post('/', upload.single('file-to-upload'), (req, res) => {
    const final_result = parseExcelResult(excelToJson({sourceFile: req.file.path}));

    if (final_result.length > 0) {
        mongoInsertResults(final_result);
    }

    res.json(final_result);
});

app.listen(3000);
