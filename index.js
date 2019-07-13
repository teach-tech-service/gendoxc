var async = require("async");
var officegen = require("officegen");
var docx = officegen("docx");
var fs = require("fs");
var path = require("path");

var table = [
  [
    {
      val: "No.",
      opts: {
        cellColWidth: 4261,
        b: false,
        color: "000000",
        sz: "48",
        shd: {
          fill: "FFFFFF",
          themeFill: "text1",
          themeFillTint: "80"
        },
        fontFamily: "Avenir Book"
      }
    },
    {
      val: "Title1",
      opts: {
        b: true,
        color: "000000",
        align: "right",
        shd: {
          fill: "FFFFFF",
          themeFill: "text1",
          themeFillTint: "80"
        }
      }
    },
    {
      val: "Title2",
      opts: {
        align: "center",
        vAlign: "center",
        cellColWidth: 42,
        b: true,
        sz: "48",
        shd: {
          fill: "FFFFFF",
          themeFill: "text1",
          themeFillTint: "80"
        }
      }
    }
  ],
  [
    [
      {
        type: "image",
        path: "",
        opts: {
          cx: 72,
          cy: 72
        }
      }
    ],
    [
      {
        type: "text",
        inline: true,
        values: [
          {
            opts: {
              b: true,
              sz: 20
            }
          },
          {
            val: " Balance Training",
            opts: {
              sz: 20
            }
          },
          {
            val: "",
            opts: {
              sz: 20
            }
          }
        ]
      },
      {
        type: "text",
        inline: true,
        values: [
          {
            opts: {
              b: true,
              sz: 20
            }
          },
          {
            val: " Beginning Knitting",
            opts: {
              sz: 20
            }
          },
          {
            val: ", Salon",
            opts: {
              sz: 20
            }
          }
        ]
      }
    ],
    "All grown-ups were once children",
    ""
  ],
  [2, "there is no harm in putting off a piece of work until another day.", ""],
  [
    3,
    "But when it is a matter of baobabs, that always means a catastrophe.",
    ""
  ],
  [4, "watch out for the baobabs!", "END"]
];

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: "ada",
  tableAlign: "left",
  tableFontFamily: "Comic Sans MS",
  borders: false
};

docx.createTable(table, tableStyle);
var out = fs.createWriteStream(path.join(__dirname, "example.docx"));

out.on("error", function(err) {
  console.log(err);
});

async.parallel(
  [
    function(done) {
      out.on("close", function() {
        console.log("Finish to create a DOCX file.");
        done(null);
      });
      docx.generate(out);
    }
  ],
  function(err) {
    if (err) {
      console.log("error: " + err);
    } // Endif.
  }
);
