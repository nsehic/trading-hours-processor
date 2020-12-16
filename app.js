const xlsxFile = require("read-excel-file/node");
const moment = require("moment");
const fetch = require("node-fetch");
const dotenv = require("dotenv");
dotenv.config();

const CURRENT_WEEK = 25;
const TIMER = 200;

async function sleep(millis) {
  return new Promise((resolve) => setTimeout(resolve, millis));
}

function formatTime(d) {
  const number = Math.round(d * 24 * 10) / 10;
  return moment(number.toString(), "LT").format("HH:mm");
}

if (process.argv[3] === undefined) {
  console.log("You must pass a filename for processing");
  process.exit();
}

const columns = {
  week: 5,
  store: 4,
  id: 1,
  omneo_id: 21,
  weekday: {
    monday: {
      open: 7,
      close: 8,
    },
    tuesday: {
      open: 9,
      close: 10,
    },
    wednesday: {
      open: 11,
      close: 12,
    },
    thursday: {
      open: 13,
      close: 14,
    },
    friday: {
      open: 15,
      close: 16,
    },
    saturday: {
      open: 17,
      close: 18,
    },
    sunday: {
      open: 19,
      close: 20,
    },
  },
};

if (process.argv[2] === "website") {
  updateWebsite(process.argv[3]);
} else if (process.argv[2] === "google") {
  console.log("This script is not supported yet");
} else {
  console.log("Unknown command");
}

function updateWebsite(filename) {
  xlsxFile(filename).then(async (rows) => {
    let count = 0;
    for (let i = 2; i < rows.length; i++) {
      await sleep(TIMER);
      if (rows[i][columns.week] === CURRENT_WEEK) {
        const omneo_id = rows[i][columns.omneo_id];
        const location = rows[i][columns.store];
        count += 1;
        const body = {
          normal_hours: [
            {
              day_of_week: "MON",
              open_at: formatTime(rows[i][columns.weekday.monday.open]),
              close_at: formatTime(rows[i][columns.weekday.monday.close]),
            },
            {
              day_of_week: "TUE",
              open_at: formatTime(rows[i][columns.weekday.tuesday.open]),
              close_at: formatTime(rows[i][columns.weekday.tuesday.close]),
            },
            {
              day_of_week: "WED",
              open_at: formatTime(rows[i][columns.weekday.wednesday.open]),
              close_at: formatTime(rows[i][columns.weekday.wednesday.close]),
            },
            {
              day_of_week: "THU",
              open_at: formatTime(rows[i][columns.weekday.thursday.open]),
              close_at: formatTime(rows[i][columns.weekday.thursday.close]),
            },
            {
              day_of_week: "FRI",
              open_at: formatTime(rows[i][columns.weekday.friday.open]),
              close_at: formatTime(rows[i][columns.weekday.friday.close]),
            },
            {
              day_of_week: "SAT",
              open_at: formatTime(rows[i][columns.weekday.saturday.open]),
              close_at: formatTime(rows[i][columns.weekday.saturday.close]),
            },
            {
              day_of_week: "SUN",
              open_at: formatTime(rows[i][columns.weekday.monday.open]),
              close_at: formatTime(rows[i][columns.weekday.sunday.close]),
            },
          ],
        };

        fetch(
          `https://api.therejectshop.omneoapp.com/api/v3/locations/${omneo_id}`,
          {
            method: "PUT",
            body: JSON.stringify(body),
            headers: {
              "Content-Type": "application/json",
              Accept: "application/json",
              Authorization: `Bearer ${process.env.BEARER_TOKEN}`,
            },
          }
        ).then((res) => {
          if (res.ok) {
            console.log(location, res.status, `(${omneo_id})`);
          } else {
            console.log(location, body);
          }
        });
      }
    }
  });
}
