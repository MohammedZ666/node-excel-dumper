require('dotenv').config()
const readXlsxFile = require('read-excel-file/node')
const path = require('path')
const fs = require('fs')
const mongoose = require('mongoose')
const excelDir = path.join(__dirname, `/data/${process.env.FILENAME}`);
const collName = process.env.COLL_NAME
const connectDB = async () => {
    try {
        const conn = await mongoose.connect(process.env.MONGO_URI, {
            useNewUrlParser: true,
            useUnifiedTopology: true,
        });
        let db = mongoose.connection.db;

        db.createCollection(collName, function (err, res) {
            if (err) {
                //console.log(err);
                if (err.codeName == "NamespaceExists") {
                    console.log("Already Exists Collection  : " + collName + "");
                    return;
                }
            }
            console.log("Collection created! : " + collName + "");

        });
        console.log(`MongoDB Connected to: ${conn.connection.host}`);
        return db;
    } catch (err) {
        console.error(err);
    }
    return null;
};
const createJsonArray = (rows) => {
    let keys = rows.shift()
    let jsonArr = [];
    keys.forEach((key, i) => {
        keys[i] = toCamelCase(key)
    });
    rows.forEach(row => {
        let json = {}
        keys.forEach((key, i) => {
            json[key] = row[i]
        });
        jsonArr.push(json)
    });
    return jsonArr;
}

const toCamelCase = (str) => {
    str = str.trim();
    var arr = str.match(/[a-z]+|\d+/gi);
    return arr.map((m, i) => {
        let low = m.toLowerCase();
        if (i != 0) {
            low = low.split('').map((s, k) => k == 0 ? s.toUpperCase() : s).join``
        }
        return low;
    }).join``;
}
const pushData = async (db, docs) => {
    try {
        await db.collection(collName).insertMany(docs);
    } catch (e) {
        console.log(e);
        throw e;
    }
}
// Readable Stream
const main = async () => {
    console.log("(node-excel-dumper) => Starting...")

    let rows = await readXlsxFile(fs.createReadStream(excelDir));
    let docs = createJsonArray(rows);
    let db = await connectDB();
    await pushData(db, docs)

    console.log("(node-excel-dumper) => Done!")
    process.exit(1)

}

main();