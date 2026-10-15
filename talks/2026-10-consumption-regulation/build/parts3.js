const fs = require("fs");
const list = ["p1", "p2", "p3", "p4", "p5", "p6", "pb"].filter((n) => fs.existsSync(__dirname + "/" + n + ".js"));
module.exports = list.map((n) => require("./" + n + ".js"));
