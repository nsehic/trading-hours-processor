const xlsxFile = require("read-excel-file/node");
const moment = require("moment");
const fetch = require("node-fetch");

const CURRENT_WEEK = 25;
const TIMER = 200;

async function sleep(millis) {
  return new Promise((resolve) => setTimeout(resolve, millis));
}

function formatTime(d) {
  const number = Math.round(d * 24 * 10) / 10;
  return moment(number.toString(), "LT").format("HH:mm");
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

xlsxFile("./files/ChristmasTradingHours.xlsx").then(async (rows) => {
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
            Authorization:
              "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6ImU2ZmFkZGI1NTc3OWJkOTk2M2I4MDgxM2UxZTFmODI3YWI0OGViYmZmNTdiMjRlMjRhZWM0OWEwN2QzZDdhNTY4ZjZmYmRiMWJjZDJlNDA3In0.eyJhdWQiOiIyIiwianRpIjoiZTZmYWRkYjU1Nzc5YmQ5OTYzYjgwODEzZTFlMWY4MjdhYjQ4ZWJiZmY1N2IyNGUyNGFlYzQ5YTA3ZDNkN2E1NjhmNmZiZGIxYmNkMmU0MDciLCJpYXQiOjE2MDc5OTEyNzksIm5iZiI6MTYwNzk5MTI3OSwiZXhwIjoxNjM5NTI3Mjc5LCJzdWIiOiIxMyIsInNjb3BlcyI6WyJjcmVhdGUtbG9jYXRpb25zIiwicmVhZC1sb2NhdGlvbnMiLCJ1cGRhdGUtbG9jYXRpb25zIiwiZGVsZXRlLWxvY2F0aW9ucyJdfQ.O4xINmt0bFmGsRmf2EuGwBxTA9OGQZGFtVbfBMvvKvIfwU3qzUkzxCnmYuiF9GX02OAsJM30JttiPZJWAkSgWvNj-0Bet6_WhK2X2_B6k_ouI46Fzk7tTxWKm1OIf9GYsqZggyD2MWAJCgTja6rbV24DegV3-eh1tWr5DdiCS1m83r1xNgpMbpbkooh7t8Fom46bBDozV2zYvVIVidu2dG5PAQ4dbsIY-Mv2iclWP45jE0-VvC7o5LtpVdMH_HBujOuK65AYqVDeHW861H2imjlcb71NqTP14gUIDG2hjQW9fTdLrZNir_EhpLI1NX9aiHDmVMA5t_tGLqW5uTwP-nvtQ2hLB3hFXLkhkYHkYjSg4wJ1aFOPCzbID1f1FXmOvdz5gkrhiCPHpgw2ncM6Kt7d4feraW9eVnpWRhwwOzdzQg2bTNzas3QpwB6VW5LQem0vV4zIpdmqYb2InZBpm7dMDYE0kKWqniNBsDdcgrftNLFgnK-dOI0P9u1tyiksRnXTvpVZveuCicT4r_0OR_T5UL3IhD8wRclvscuWQmxbl1_J3XbcLL8F7rJU80rLPXkH_b-AjBtOv9UHqUc7UEga-oaZ4SAkuYRKegQhl-30Q3PdeNwnD1jHkxmVFs-id_ciqNtcYhV3fjaOswB9LXTyARPI7VPqJaXnl6k7FqY",
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
