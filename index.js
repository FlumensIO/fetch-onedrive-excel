// get local environment variables from .env
const http = require("https"); // eslint-disable-line
const exec = require("child_process").exec;
const readline = require("readline");

async function fetch({ drive = "me/drive", file, sheet }) {
  console.log(`Pulling ${sheet} from ${file}.`);

  let token = process.env.MS_TOKEN;
  if (!token) {
    const tokenMessage = `\n\n🥺 Requires an MS_TOKEN env var set up. You can get it using this link:\n\n  ↪ https://developer.microsoft.com/en-us/graph/graph-explorer/preview\n\n`;

    var isMacOS = process.platform === "darwin";
    if (!isMacOS) {
      return Promise.reject(new Error(tokenMessage));
    }

    console.warn(tokenMessage);

    // open browser
    exec('open "https://developer.microsoft.com/en-us/graph/graph-explorer"');

    // accept the token
    const prompt = new Promise((resolve) => {
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
      });

      rl.question("Please enter the Access token here:\n", function (value) {
        resolve(value);
        rl.close();
      });
    });

    token = await prompt;

    process.env.MS_TOKEN = token; // saving for other script runs in the same process
  }

  if (!drive) {
    throw new Error("Drive option is missing.");
  }

  if (!file) {
    throw new Error("File option is missing.");
  }

  if (!sheet) {
    throw new Error("Sheet option is missing.");
  }

  return new Promise((resolve) => {
    const options = {
      method: "GET",
      hostname: "graph.microsoft.com",

      // examples:
      // /workbook/worksheets('species')/usedRange
      // /workbook/worksheets('species')/range(address='species!A1:FC73')
      // /workbook/worksheets('species')/cell(row=3,column=8)
      path: `/v1.0/${drive}/items/${file}/workbook/worksheets('${sheet}')/usedRange`,
      headers: {
        Authorization: `Bearer ${token}`,
        "Cache-Control": "no-cache",
      },
    };

    const req = http.request(options, (res) => {
      const chunks = [];

      res.on("data", (chunk) => {
        chunks.push(chunk);
      });

      res.on("end", () => {
        const body = Buffer.concat(chunks);
        const json = JSON.parse(body.toString());

        if (json.error) {
          throw new Error(json.error.message || "Request error");
        }

        resolve(json.values);
      });
    });

    req.end();
  });
}

function setColumnObj(obj, columnName, value) {
  if (!columnName.length) {
    return null;
  }

  // cleanup
  const keyFull = columnName.replace(" ", "");

  // check if any fancy parsing is needed
  const array = keyFull[keyFull.length - 1] === "]";
  const object = keyFull[keyFull.length - 1] === "}";

  if (!array && !object) {
    const isEmptyValue =
      typeof value === "string" && !value.replaceAll(" ", "").length;
    if (isEmptyValue) {
      return;
    }

    return (obj[columnName] = value); // eslint-disable-line
  }

  let index;
  let key;
  if (array) {
    index = keyFull.indexOf("[");
    key = keyFull.substring(0, index);

    // initialize if non existing
    if (!obj[key]) {
      obj[key] = []; // eslint-disable-line
    }

    return obj[key].push(value);
  }

  if (object) {
    index = keyFull.indexOf("{");
    key = keyFull.substring(0, index);

    // initialize if non existing
    if (!obj[key]) {
      obj[key] = {}; // eslint-disable-line
    }

    const newColumnName = keyFull.substring(index + 1, keyFull.length - 1);
    return setColumnObj(obj[key], newColumnName, value);
  }
  return null;
}

function normalizeValue(value) {
  // check if int
  // https://coderwall.com/p/5tlhmw/converting-strings-to-number-in-javascript-pitfalls
  const int = value; // * 1;
  if (!Number.isNaN(int)) return int;
  return value;
}

function processRow(header, row) {
  const rowObj = {};

  // for each row column value set it in the right rowObj
  row.forEach((columnVal, i) => {
    if (columnVal !== null && columnVal !== "" && columnVal !== "\n") {
      setColumnObj(rowObj, header[i], normalizeValue(columnVal));
    }
  });

  return rowObj;
}

function processRows(rows) {
  console.log("Processing rows.");

  const obj = [];
  const header = rows.shift();
  rows.forEach((row) => {
    const rowObj = processRow(header, row);
    obj.push(rowObj);
  });

  return obj;
}

module.exports = async (fileID, sheetName) => {
  const data = await fetch(fileID, sheetName);
  return processRows(data);
};
