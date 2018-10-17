var express = require('express');
var fs = require('fs');
var path = require('path');
var async = require ( 'async' );
var officegen = require('officegen');
var style = require('../lib/style');
var CircularJSON = require('circular-json')

var router = express.Router();

var docx =null;
var document = {
  name : "",
  type : "docx",
  creator :"",
  title : "",
  description : "",
  author : "",
  subject : "",
  keywords : "",
  orientation:"",
  pageMargins:""
};



/* Document Properties. */
router.route('/')
.get( function(req, res, next) {
  if(document.name !== ""){
    var out = fs.createWriteStream ( './output/'+document.name+'.docx' );
    docx.generate ( out );
    out.on ( 'close', function () {
      console.log ( 'Finish to create a DOCX file.' );
      res.status(201).json(JSON.parse(CircularJSON.stringify(docx)));
    });

  }
})
.put( function (req, res, body){
  document = req.body;
  docx = officegen(document);

  res.json(document);
})
.post( function (req, res, body){
  res.status(201).json(JSON.parse(CircularJSON.stringify(docx)));
});







/** Document Table of content */
router.route('/tabContents')
.get( function (req, res, body){
  res.send('Getting the Docuemnt table of Contents to implement');
})
.post( function (req, res, body){


  var pObj = docx.createP ();
  pObj.addText ( "Table of Contents", style.Heading1);
  pObj.addHorizontalLine ();

  res.send('Getting the Docuemnt header to implement');
})









/** Document Cover Page */
router.route('/cover')
.get( function (req, res, body){
  res.send('Getting the Docuemnt header to implement');
})
.post( function (req, res, body){

  /** Aggiunta immagine Logo */
  var pObj = docx.createP ();
  pObj.options.align = 'left';
  console.log("Image : ",req.body.image);
  pObj.addImage ( path.resolve(__dirname, req.body.image), { cx: 150, cy : 50 } );

  /** Spaziatura per titolo */
  for(var i=0; i<req.body.blankLines; i++)
      pObj.addLineBreak ();

  /** Aggiunta Titolo copertina */
  pObj.addText ( req.body.title, style.Title);
  pObj.addLineBreak ();
  pObj.addHorizontalLine ();

  pObj.addText ( req.body.subTitle, style.Subtitle);
  docx.putPageBreak ();


  res.status(200).json({message:"Image added to header "});

})
.put( function (req, res, body){
  res.send('Modify the document header to implement');
})
.delete( function (req, res, body){
  res.send('Delete the document header to implement');
})





/** Document Contact Information */
router.route('/contact')
.get( function (req, res, body){
  res.send('Getting the Docuemnt header to implement');
})
.post( function (req, res, body){

    var pObj = docx.createP ();
    pObj.addText ( req.body.contactTitle, style.Heading1);
    pObj.addHorizontalLine ();

    docx.createTable (req.body.data, req.body.style);
    pObj.addLineBreak();



    var pObj = docx.createP ();
    pObj.addText ( req.body.confidentialityTitle, style.Subtitle);
    pObj.addLineBreak();

    var pObj = docx.createP ();
    pObj.options.align = "justify";
    pObj.addText ( req.body.confidentialityBody, style.Body);
    docx.putPageBreak ();


  res.send('Getting the Docuemnt header to implement');
})
.put( function (req, res, body){
  res.send('Putting the Docuemnt contact to implement');
})
.delete( function (req, res, body){
  res.send('deleting the Docuemnt contact to implement');
});





/** Document General Information */
router.route('/docInfo')
.get( function (req, res, body){
  res.send('Getting the docInfo  to implement');
})
.post( function (req, res, body){

  var pObj = docx.createP ();
  pObj.addText ( req.body.docInfoTitle, style.Heading1);
  pObj.addHorizontalLine ();
  pObj.addLineBreak();

  var pObj = docx.createP ();
  pObj.addText ( req.body.docInfoSec1, style.Heading2);
  docx.createTable (req.body.dataSec1, req.body.style);

  pObj.addLineBreak();
  pObj.addLineBreak();

  
  var pObj = docx.createP ();
  pObj.addLineBreak();
  pObj.addText ( req.body.docInfoSec2, style.Heading2);
  pObj.addLineBreak();
  if(req.body.dataSec2.length >0) {
    console.log("Creazione Tabella. Numero elementi: ",req.body.dataSec2.length); 
    docx.createTable (req.body.dataSec2, req.body.style);
  } else {
    console.log("Tabella vuota. Esco con messaggio di warning");
    pObj.addText ( req.body.docInfoSecMissing, style.body);
  }

  
  pObj.addLineBreak();
  pObj.addLineBreak();

  var pObj = docx.createP ();
  pObj.addText ( req.body.docInfoSec3, style.Heading2);
  pObj.addLineBreak();
  if(req.body.dataSec3 >0){
    docx.createTable (req.body.dataSec3, req.body.style);
  } else {
    pObj.addText ( req.body.docInfoSecMissing, style.body);
  }
  


  docx.putPageBreak ();
  res.send('Posting the docInfo header to implement');
});






router.route('/genericParagraph')
.get( function (req, res, body){
  res.send('Getting the genericTitle  to implement');
})
.post( function (req, res, body){
  var pObj = docx.createP ();
  pObj.addText ( req.body.docTitle, req.body.style);

  for(var i=0; i<req.body.horizontalLine; i++)
    pObj.addHorizontalLine ();

  for(var i=0; i<req.body.lineBreak; i++)
    pObj.addLineBreak();

  for(var i=0; i<req.body.pageBreak; i++)
    docx.putPageBreak ();


  res.send('generic Paragraph section has been added');
})


router.route('/genericTable')
.get( function (req, res, body){
  res.send('Getting the genericTable  to implement');
})
.post( function (req, res, body){
  docx.createTable (req.body.data, req.body.style);


  res.send('generic Paragraph section has been added');
})


module.exports = router;
