const fs = require("fs");
const readline = require("readline");
const { google } = require("googleapis");
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"];
const express = require("express");
const app = express();
const port = 3000;
let oAuthParam = "";

const TOKEN_PATH = "token.json";
fs.readFile("./credentials.json", (err, content) => {
  if (err) return console.log("Error loading client secret file:", err);
  authorize(JSON.parse(content));
});

function authorize(credentials) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client);
    oAuth2Client.setCredentials(JSON.parse(token));
    oAuthParam = oAuth2Client;
  });
}

function getNewToken(oAuth2Client) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES
  });
  console.log("Authorize this app by visiting this url:", authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question("Enter the code from that page here: ", code => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err)
        return console.error(
          "Error while trying to retrieve access token",
          err
        );
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), err => {
        if (err) return console.error(err);
        console.log("Token stored to", TOKEN_PATH);
      });
      oAuthParam = oAuth2Client;
    });
  });
}

getStatus = x => {
  if (x.red == 1) return 1;
  if (x.green == 1 && x.red == undefined) return 2;
  return 0;
};

function listMajors(auth, startIndex, endIndex) {
  return new Promise(function(resolve) {
    const sheets = google.sheets({ version: "v4", auth });
    let ranges = [
      `Fixed Schedule-Term IV!D${startIndex}:D${endIndex}`,
      `Fixed Schedule-Term IV!E${startIndex}:E${endIndex}`,
      `Fixed Schedule-Term IV!F${startIndex}:F${endIndex}`,
      `Fixed Schedule-Term IV!G${startIndex}:G${endIndex}`,
      `Fixed Schedule-Term IV!H${startIndex}:H${endIndex}`,
      `Fixed Schedule-Term IV!I${startIndex}:I${endIndex}`
    ];
    sheets.spreadsheets.get(
      {
        spreadsheetId: "140aP6lqCZuo_NalN5cVIMia0zI5rISzriuN88fIbm94",
        includeGridData: true,
        ranges: ranges
      },
      (err, res) => {
        let mySubjects = {};
        datas = res.data.sheets[0].data;
        for (let d = 0; d < datas.length; d++) {
          let rowData = datas[d].rowData;
          for (let i = 0; i < rowData.length; i++) {
            if (rowData[i].values) {
              subj = rowData[i].values[0].userEnteredValue;
              if (subj && subj.stringValue != undefined) {
                if (mySubjects[subj.stringValue] == undefined) {
                  mySubjects[subj.stringValue] = [
                    {
                      classRoom: d,
                      timing: i,
                      status: getStatus(
                        rowData[i].values[0].userEnteredFormat.backgroundColor
                      )
                    }
                  ];
                } else {
                  let t = mySubjects[subj.stringValue];
                  t = t.concat({
                    classRoom: d,
                    timing: i,
                    status: getStatus(
                      rowData[i].values[0].userEnteredFormat.backgroundColor
                    )
                  });
                  mySubjects[subj.stringValue] = t;
                }
              }
            } else console.log("NOT A PERIOD");
          }
        }
        if (err) return console.log("The API returned an error: " + err);
        resolve(mySubjects);
      }
    );
  });
}

app.get("/", (req, res) => {
  let index = req.query.index;
  startIndex = 608 + index * 11;
  endIndex = 618 + index * 11;
  let items = listMajors(oAuthParam, startIndex, endIndex);
  items.then(result => {
    console.log(result);
    res.send(result);
  });
});

app.listen(port, () => console.log(`Example app listening on port ${port}!`));
