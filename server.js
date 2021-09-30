const Express = require("express");
const Cors = require("cors");
const BodyParser = require("body-parser");
const MongoClient = require("mongodb").MongoClient;
const { google } = require("googleapis");

const CONNECTION_URL =
  "mongodb+srv://shay:12345@cluster1.dlgr8.mongodb.net/exp?retryWrites=true&w=majority";
const DATABASE_NAME = "exp";

var app = Express();
app.use(Cors());
app.use(BodyParser.json());
app.use(BodyParser.urlencoded({ extended: true }));
var database, collection;

//connection start
app.listen(5000, () => {
  MongoClient.connect(CONNECTION_URL, (error, client) => {
    if (error) {
      throw error;
    }
    database = client.db(DATABASE_NAME);
    collection = database.collection("data");
    console.log("Connected to `" + DATABASE_NAME + "`!");
  });
});

const auth = new google.auth.GoogleAuth({
  keyFile: "./config.json", //the key file
  //url to spreadsheets API
  scopes: "https://www.googleapis.com/auth/spreadsheets",
});

const spreadsheetId = "175TepXJJF3-nvt7xuafSMN5TnVngtcGxHSU-L5VuXJc";

// export data
app.post("/exportdata", (request, response) => {
  //   console.log(request.body);

  const query = { fname: request.body.fname };
  const update = {
    $set: request.body,
  };
  const options = { returnNewDocument: true };
  return collection
    .findOneAndUpdate(query, update, options)
    .then((updatedDocument) => {
      //   console.log(updatedDocument.lastErrorObject);
      if (updatedDocument.lastErrorObject.updatedExisting) {
        gsheetfind(request.body.fname, [
          request.body.fname,
          request.body.lname,
          request.body.email,
          request.body.phone,
        ]);
        // console.log(`Successfully updated document: ${updatedDocument}.`);
        // console.log(updatedDocument.value)
        response.json(updatedDocument);
      } else {
        console.log("No document matches the provided query.");
        collection.insertOne(request.body, (error, result) => {
          if (error) {
            return response.status(500).send(error);
          }
          console.log("Data Sent", result);
          gsheet([
            request.body.fname,
            request.body.lname,
            request.body.email,
            request.body.phone,
          ]);
          response.json(result);
        });
      }
      return updatedDocument;
    })
    .catch((err) => {
      console.error(`Failed to find and update document: ${err}`);
    });
});

async function gsheet(data) {
  console.log("inside gsheet", data);
  const authClientObject = await auth.getClient();
  const googleSheetsInstance = await google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  await googleSheetsInstance.spreadsheets.values.append(
    {
      auth, //auth object
      spreadsheetId, //spreadsheet id
      range: "Sheet1!A:B", //sheet name and range of cells
      valueInputOption: "USER_ENTERED", // The information will be passed according to what the usere passes in as date, number or text
      resource: {
        values: [data],
      },
    },
    async (err, result) => {
      if (err) {
        console.log(err);
      } else {
        console.log(result.data);
      }
    }
  );
}

async function gsheetfind(data, updatedData) {
  const authClientObject = await auth.getClient();
  const googleSheetsInstance = await google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  await googleSheetsInstance.spreadsheets.values.get(
    {
      auth, //auth object
      spreadsheetId, //spreadsheet id
      range: "Sheet1!A:A", //sheet name and range of cells
    },
    async (err, result) => {
      if (err) {
        console.log(err);
      } else {
        var rangeToUpdate = "";
        // console.log(result.data.values);
        const numRows = result.data.values ? result.data.values.length : 0;
        console.log(`${numRows} rows retrieved.`);
        var range = 0;
        result.data.values.find((name) => {
          if (name[0] === data) {
            rangeToUpdate = `Sheet1!A${range + 1}`;
            console.log("Range  to Update: ", rangeToUpdate);
          }
          range++;
        });
        await googleSheetsInstance.spreadsheets.values.update(
          {
            auth, //auth object
            spreadsheetId, //spreadsheet id
            range: rangeToUpdate, //sheet name and range of cells
            valueInputOption: "USER_ENTERED", // The information will be passed according to what the usere passes in as date, number or text
            resource: {
              values: [updatedData],
            },
          },
          (err, result) => {
            if (err) {
              console.log(err);
            } else {
              console.log(result.data);
            }
          }
        );
      }
    }
  );
}
