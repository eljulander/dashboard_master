/*
 *   DASHBOARD MASTER
 *
 *   This is the master script for all of the course dashboards.
 *   
 *   This script has four main functions:
 *   1.) Pushes data back to the "Conversion Progress" spreadsheet.
 *   2.) Pushes the data to the Firebase server
 *   3.) Changes spreadsheet theme based off of the current team
 *   4.) Sets the template for themes
 */

/*
 * Creates a menu option for the spreadsheet themes, 
 * and then updates spreadsheet data when the spreadsheet is first open. 
 */
function onOpen(){
   //creates menu
  if(!validateUser())return;
  var menu = SpreadsheetApp.getUi();
   menu.createMenu("Style")
  .addItem("Style Me", "setFormat")
  .addItem("Set Me", "getFormat")
  .addToUi();
  
   sendDataToConversionProgress();
   sendDataToFirebase();

}



/*
 * This finds if the current user is authorized to use the current spreadsheet
 */

var authUsers = [
  "ericjulander@gmail.com","mooreco.byui@gmail.com","johnsonca@byui.edu","danniellehext11@gmail.com","egfischvogt@gmail.com","mckenzier.byui@gmail.com","jade.coates@gmail.com","andrew.gremlich@gmail.com","wubblebee@gmail.com","HaileyAndrick@gmail.com","gwinert@gmail.com","halesm.byui@gmail.com"];

function validateUser(){
  var userFound = false;
  var user;
  try{
    user = Session.getActiveUser().getEmail();
  }catch(e){
    return true;	
  }
  for(var i in authUsers)
    if(authUsers[i] == user)
    {
      userFound = true;
      break;
    }
  
  if(!userFound)
    alert("Sorry, you dont have authorization to press this button.");
  
  return userFound;
}



/*
 * Pushes information from this spreadsheet to the parent sheet
 */
function pushToMotherSheet(){
  try{
  ScriptMaster.pushToMotherSheet();
  }catch(e){Logger.log(e);}
  var copied = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("V2").getValue();
  onOpen();
  if(!copied)
  {
    setFormat();
    SpreadsheetApp.getActive().getSheetByName("Variables").getRange("V2").setValue("TRUE");
  }
}

// style default variables
var format = DocumentApp.openById("1CcCK0mCuaTRkfgdY5iZQr_V9DYYRa_0cM2S90NdJT7Y").getBody();
var ss = SpreadsheetApp.getActive();
var json = '{"colors":{"gryffindor":{"font1":"","font2":"","font3":"","color1":"#8f1609","color2":"#885d00","color3":"#0b6188","border":""},"ravenclaw":{"font1":"","font2":"","font3":"","color1":"#054a6f","color2":"#5b4937","color3":"","border":""},"slytherin":{"font1":"","font2":"","font3":"","color1":"#2e4608","color2":"#545d54","color3":"","border":""},"hufflepuff":{"font1":"","font2":"","font3":"","color1":"#d19405","color2":"#1a1a1a","color3":"","border":""},"template":{"font1":"","font2":"","font3":"","color1":"#054a6f","color2":"#5b4937","color3":"#0b6188","border":""}},';


// sets template colors for the different teams
var colors = {
        gryffindor:{
           color1: "#8f1609",
           color2: "#885d00",
           color3: "#0b6188"
        },
        ravenclaw:{
        
           color1: "#054a6f",
           color2: "#5b4937",
           color3: ""
        
        },
        slytherin:{
           color1: "#2e4608",
           color2: "#545d54",
           color3: "",
           border: ""
        },
        hufflepuff:{
           color1: "#d19405",
           color2: "#1a1a1a",
           color3: ""
        }
    
};

// colors that define the template
var templateColor = ["#054a6f","#5b4937"];


var border = "";
var header = "";
var table = {
  
  tableformat:{

  }

};


function getFormat(){
	getzFormat();
}

/*
 *
 */
function getzFormat(){
  for(var s in ss.getSheets()){
    header = "";
    border = "";
    var r = ss.getSheets()[s].getDataRange().getBackgrounds();
    for(var y in r){
      for(var x in r[y]){
        var A1 = String.fromCharCode(parseInt(x)+65) + (parseInt(y)+1);
        addToJSON(A1 + ":"+r[y][x]);
      }
    }
    
    Logger.log(ss.getSheets()[s].getName());    
    addToTemplate(header,border,ss.getSheets()[s].getName());
    
  }
  
  
  format.setText(json+tableformat+"}}");
  
}

var tableformat = '"tableformat":{';
function addToTemplate(header, border, name)
{
  var t = '"'+name+'":{'
  var b = '"border":{"range":['+border+'""],"color":"color1"},';
  var h = '"header":{"range":['+header+'""],"color":"color2"},';
  t += b;
  t += h;
  t += "},";
  tableformat += t;
}

function addToJSON(color){
  
  var A1 = color.split(":")[0];
  color = color.split(":")[1];
  
  if(color == templateColor[0]){
     border += '"'+A1+'"' + ",";
  
  }
  else if(color == templateColor[1])
  {
     header += '"'+A1+'"' + ",";
  }
 
  
}

function setFormat(){
  formatMe();
}


var formatData;

/*
 * Gets the template to style spreadsheet
 */
function getFormatData(){
  var json = DocumentApp.openById("1CcCK0mCuaTRkfgdY5iZQr_V9DYYRa_0cM2S90NdJT7Y").getBody().getText();
  formatData = JSON.parse(json);
}

/*
 * Styles spreadsheet according to the template
 */
function formatMe() {
  
  getFormatData();
  
  var house = getHouse();
  var data  = formatData.tableformat;
 
  for(var b in data){
    Logger.log(b);
    for(var a in data[b]){
      for(var c in data[b][a].range){
        try{ 
         SpreadsheetApp.getActiveSpreadsheet().getSheetByName(b).getRange(data[b][a].range[c]).setBackground(formatData.colors[house][data[b][a].color]);
        }catch(exception){Logger.log(exception);}
      }
  }
 }
}

/*
 * Searches spreadsheet to find proper style
 */
function getHouse(){
  var index = indexer2("Team");
  switch(SpreadsheetApp.getActive().getSheets()[0].getRange(index[1]+1, index[0]).getValue().trim().substr(0,1).toLowerCase()){
    case 'g':
      return 'gryffindor';
      break;
    case 's':
      return 'slytherin';
      break;
    case 'r':
      return 'ravenclaw';
      break;
    case 'h':
      return 'hufflepuff';
      break;
    case 't':
      return 'template';
      break;
   
  }
  return 'error';
}


/*
 * This gets data from the dashboard
 */
function indexer2(name){
  var index = -2;
  var row = -2;
  var sheet = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getValues();
  for(var b in sheet){
    for(var a in sheet[0]){
      if(sheet[b][a] == name){
        index = a;
        row = b;
        break;
      }
    }
  }
  return [parseInt(index)+1, parseInt(row)+1];
}


/*
 * Sends data back to the conversion progress sheet
 */
function sendDataToConversionProgress(){
  
  var spreadsheetData = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("W1");
  spreadsheetData.setFormulaR1C1('=IMPORTRANGE("1p43wwtkbT0K0QUpDtvqKjdZ8mgdojNP0Xu4z4gX2RYE","pensive!A:A")');
  spreadsheetData = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("W:W").getValues();
  var key = spreadsheetData[0][0];
 
  var currentCourse = SpreadsheetApp.getActive().getName().split("Dashboard")[0].trim();
  
  var oldSheet = SpreadsheetApp.openById(key).getSheets()[0];
  var names = oldSheet.getRange("A:A").getValues();
  var currentRow = 305;
  
  for(var i in names)
    if(currentCourse == names[i][0]){
      currentRow = parseInt(i)+1;
      Logger.log(currentRow);
      break;
   }
  
  var first = true;
  for(var x in spreadsheetData){
    if(spreadsheetData[x] == "")break;

    try{
      Logger.log(x);
      var location = spreadsheetData[x][0].split(":")[1];
      var index = indexer2(spreadsheetData[x][0].split(":")[0]);
   
      var currentData = (SpreadsheetApp.getActive().getSheets()[0].getRange(index[1]+1,index[0]).getValue());
      var currentNotation = location+currentRow;
      Logger.log("Sending value: "+currentData+" to "+currentNotation);
      oldSheet.getRange(currentNotation).setValue(currentData);
    }catch(e){
      
    }
  }
}
  


  
/*
 * Creates a ppup in the spreadsheet window
 */  
function alert(a){
  SpreadsheetApp.getUi().alert(a);
}

/*
 * Gets the course name
 */
function getName(){
  return SpreadsheetApp.getActive().getName().split("Dashboard")[0].trim();
}

/*
 * Converts Simplified team name to expanded team name
 */
function getTeam(){
  var currentTeam = SpreadsheetApp.getActive().getSheets()[0].getRange("H6").getValue();
  
  var names = {
    "GT 01":"GryffindorOne",
    "GT 02":"GryffindorTwo",
    "GT 03":"GryffindorThree",
    
    "ST 01":"SlytherinOne",
    "ST 02":"SlytherinTwo",
    "ST 03":"SlytherinThree",
    
    "HT 01":"HufflepuffOne",
    "HT 02":"HufflepuffTwo",
    "HT 03":"HufflepuffThree",
    
    "RT 01":"RavenclawOne",
    "RT 02":"RavenclawTwo",
    "RT 03":"RavenclawThree",
  };
  
  return names[currentTeam];
  
}

// Firebase data structure
var def = {
  "details" : {
    "dashboard" : SpreadsheetApp.getActive().getUrl(),
    "status" : "",
    "week" : ""
  },
  "dir" : {
    "cd" : {
      "email" : "",
      "name" : ""
    },
    "cl" : {
      "email" : "",
      "name" : ""
    },
    "sl" : {
      "name" : ""
    },
    "tl" : {
      "name" : ""
    }
  },
  "links" : {
    "il2" : "",
    "il3" : ""
  },
  "progress" : {
    "phase" : ""
  }
};

/*
 * Sends the dashboard data to the Firebase Server
 */
function sendDataToFirebase(){
  var currentURL = "https://mmap.firebaseio.com/";
  var dataLink = getTeam()+"/"+getName();
  
  currentURL += dataLink;
  Logger.log(currentURL);
  
  var database = getDatabaseByUrl(currentURL, "rs4XlisWL1eJNTdWkuUx9LrYdH8b8xilyerrJz5B");
  
  var data = database.getData() || def;

  
  database.setData("",addData(data));
  
}

/*
 * Adds data to the correct Firebase directory
 */
function addData(data){
  
  var refs = {
    //"B3":"links/il2",
    //"C3":"links/il3",
    "D3":"details/status",
    "E3":"progress/phase",
    "G3":"dir/sl/name",
    "H3":"dir/tl/name",
    "B6":"dir/cl/name",
    "D6":"dir/cl/email",
    "E6":"dir/cd/name",
    "F6":"dir/cd/email",
    "G6":"details/week"
  }; 
  Logger.log(data);
  for(var dir in refs){
    var currentValue = SpreadsheetApp.getActive().getSheets()[0].getRange(dir).getValue();
    Logger.log(currentValue);
    var locations = refs[dir].split("/");
    if(locations.length == 2)
    data[locations[0]][locations[1]] = currentValue;
    else
     data[locations[0]][locations[1]][locations[2]] = currentValue;
  }
  
  return data;
  
}


