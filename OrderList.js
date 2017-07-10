/**
* 
* @author Mehul Shinde
* this script generates an order
*/
var orderList=new Array();//in column 0 is name of the item, column 1 has quantity
var itemArray=new Array();
var date;//the start date of order-cycle
var startRow=1;//row of start date
var message="Please enter the start date of your order in the format <month name> <date> <year> (Ex. May 5 2017)";
var doc = DocumentApp.openById('1gEP2M5LNAjJGyCS8YUYcKRwsvCkkVSukcGfXvAiFNTA');//Change this to a document on google drive
var ss = SpreadsheetApp.openByUrl("https://docs.google.com/a/iastate.edu/spreadsheets/d/1hTkHMmANz7DqmRXeEPiIJj843nelFUec-4OxLg2EGfk/edit?usp=sharing");//Change this to original sheet url
SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
var data= ss.getDataRange().getValues();
/**
* Adds a new item to the list
* @param takes in item name to be added  to the order list
*/
function addItem(itemName)
{
orderList.push([itemName,1]);
itemArray.push(itemName);
}
/**
* Scans the response spreadsheet for orders, adds new items or appends the quantity
* of already exisiting item in the list
*/
function scanForOrders()
{
//This part checks if the inputdate is present, and if it's present it starts scanning
if(checkStartDate())
{
for(var i=startRow; i<data.length;i++)
{
for(var j=4;j<data[0].length;j++)
{
//12,18,32 columns are for comments and don't contain order items
if(data[i][j] && (j!=12 && j!=18 && j!=32))//if the current element isn't null
{
var index=binarySearch(data[i][j],itemArray);
if(index==-1)
addItem(data[i][j])
else
orderList[index][1]+=1;
}

}
}
generateList();
}
else
inputBox();

}
/**
* Checks whether the entered date exists in the sheet
*/
function checkStartDate()
{
for(var k=1;k<data.length;k++)
{

if(data[k][0].toString().indexOf(date)!=-1)
{
startRow=k;
return true;
}
else if(k==data.length-1)
{
message="Invalid date, please enter a valid date in the format <month name> <date> <year> (Ex. May 5 2017)";
return false;
}
}

}
/**
* Generates an order list
*/
function generateList()
{
SpreadsheetApp.setActiveSheet(ss.getSheets()[1]);
var ts=ss.getActiveSheet();
ts.clear();
ts.getRange(1, 1,orderList.length,orderList[0].length).setValues(orderList);//ss.getLastRow()+1
}
/**
* Set required trigger to this
*/
function main()
{
inputBox();
//SpreadsheetApp.getUi().alert("Your order list is ready!");
}
/**
* Implements binary search and returns the position the element is at. 
* if not found, returns -1
* @param {element to search, array} 
* @returns {number} position if found, -1 if not found 
*/
function binarySearch(searchElement, searchArray) {
    'use strict';
    searchArray.sort();
    var stop = searchArray.length;
    var last, p = 0,
        delta = 0;

    do {
        last = p;

        if (searchArray[p] > searchElement) {
            stop = p + 1;
            p -= delta;
        } else if (searchArray[p] === searchElement) {
            // FOUND A MATCH!
            return p;
        }

        delta = Math.floor((stop - p) / 2);
        p += delta; //if delta = 0, p is not modified and loop exits

    }while (last !== p);

    return -1; //nothing found

}
/**
* The box to enter date and select functinality 
*/
function inputBox()
{
// Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
 // user can also close the dialog by clicking the close button in its title bar.
 var ui = SpreadsheetApp.getUi();
 
 var response = ui.prompt('Welcome!', message, ui.ButtonSet.OK_CANCEL);
 // Process the user's response.
 var response2 = ui.prompt('Welcome!', "Do you also want to individual order lists?", ui.ButtonSet.YES_NO);
 if (response.getSelectedButton() == ui.Button.OK) {
 
    date=response.getResponseText();
    scanForOrders();
   
   
 }
if(response2.getSelectedButton()==ui.Button.YES)
{
date=response2.getResponseText();
generateDocs();
}
else if (response.getSelectedButton() == ui.Button.CANCEL) {
   Logger.log('User cancelled the process');
 } else {
   Logger.log('The user clicked the close button in the dialog\'s title bar.');
 }
}
/**
* Generates individual order documents on google drive
*/
function generateDocs()
{
if(checkStartDate())
{
// Open a document by ID.
var counter=1;
var items=new Array();
var finalList=new Array();
clearBodyFunction();
var body = doc.getBody();
//body.clear();
var text = body.editAsText();
//text.insertText(0, 'Sprout\n');
for(var i=startRow; i<data.length; i++)
{
 text.appendText("Sprout\n");
 text.appendText(data[i][2]+'\n');
 text.appendText(data[i][1]+'\n');
 text.appendText(data[i][3]+'\n');//4, 13, 19 are
 var arrCount=0;
 if(data[i][4])
 {
 for(var x=4; x<=11; x++)//copy items to array
 {
 items.push(data[i][x]);
 //body.appendListItem(data[i][x]).setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);counter++;
 }
 }
 else if(data[i][13])
 {
 for(var x=13; x<=17; x++){
 items.push(data[i][x]);
 //body.appendListItem(data[i][x]);
 }
 }
 
 else if(data[i][19])
 {
 for(var x=19; x<=31; x++){
 items.push(data[i][x]);
 //body.appendListItem(data[i][x]);
 }
 }
  // finalList.push([items[0],1]);//add first item to final list
 for(var y=0;y<items.length;y++)
 {
 var ind=binarySearch(items[y],finalList);//look for each item
 if (ind==-1)
 finalList.push([items[y],1]);
 else
 finalList[ind][1]+=1;
 }
 for(var f=0; f<finalList.length;f++)
 {
 body.appendListItem(finalList[f][0]+"   x "+finalList[f][1]).setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
 Logger.log("%s x %s",finalList[f][0],finalList[f][1]);
 }
 finalList=new Array();
 items=new Array();
 body.appendPageBreak();
text.appendText('\n\r');


}
}
else
inputBox();
}
/**
* Clears the document
*/
function clearBodyFunction()
{
doc.getBody().clear();
}



