let fs = require("fs");
let officegen = require("officegen");

let jsonContent = fs.readFileSync("./Data.json");
let jsonData = JSON.parse(jsonContent);
console.log("--- JSON Data ---");
console.log(jsonData);

let docx = officegen({
	'type' : 'docx',
	'subject' : 'JSONData',
	'description' : 'Generated document using officegen node module'
});

let pObj = docx.createP();
pObj.options.align = 'left'; // Also 'right' or 'jestify'.

for(let data in jsonData){
	pObj.addText ( 'key: ' + data, {bold: true} );
	pObj.addLineBreak ();
	pObj.addText ( 'value: ' + jsonData[data] );
	pObj.addLineBreak ();
}

var out = fs.createWriteStream ( 'out.docx' );
 
docx.generate ( out );
out.on ( 'close', function () {
  console.log ( 'Finished to create the docx file!' );
});
 
docx.on ( 'error', function ( err ) {
      console.log ( err );
});