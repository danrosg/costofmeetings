var item;
var paygrades = {}; //hashmap for the paygrades
var userid_grades = {};
var recipients=[];


(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    item = Office.context.mailbox.item;
    jQuery(document).ready(function(){

    //loadCupsOfCoffee();

      // We will preload the files in memory during the first load of the page
      $.ajax({
          type: "GET",
          url: "../../database/paygrades.csv",
          dataType: "text",
          success: function(data) {processData(data,paygrades);}
       });

       $.ajax({
           type: "GET",
           url: "../../database/userid_grades.csv",
           dataType: "text",
           success: function(data) {processData(data,userid_grades);}
        });

      $('#insert-button').on('click', function(){
        //printData(paygrades);
        //printData(userid_grades);
        getAllRecipients();
        buildCoffeeList("#coffee-list",10);

      })

    });
  };

})();

//funtion to debug the csv load
function printData(map){

  write (map['DROSALES']+'\n');
  write (map[1]+'\n');
}


function processData(allText,map) {
    var allTextLines = allText.split(/\r\n|\n/);
    for( var i=1; i<allTextLines.length; i++)
    {
        var data = allTextLines[i].split(',');
        if(data.length=2)
        {
          map[data[0]] = data[1];

        }


    }

}


// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    // Get any to-recipeints.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients.
            //write ('To-recipients of the item: ');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            //write ('Cc-recipients of the item: ');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Display the email addresses of the bcc-recipients.
          //  write ('Bcc-recipients of the item: ');
            displayAddresses(asyncResult);
        }

        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {

    var total=0;

    for (var i=0; i<asyncResult.value.length; i++)
    {

        var data =asyncResult.value[i].emailAddress;
        var userid = data.split('@');
        var name = userid.length==2 ? userid[0].toUpperCase() : null;

        //
        recipients.push(name);

        total = total + userid_grades[name];
        write ( name+' '+total+'\n');

    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message ;
}




function buildCoffeeList(parent,quantity) {

  var i;

  for (i=0;i<quantity;i++){

    var listItem = $('<li/>')
      .attr("hidden","hidden")
      .appendTo(parent);

    var desc = $('<img/>')
      .attr("src","../../assets/coffee-icon-small.png" )
      .attr("alt","cup of coffee")
      .appendTo(listItem);

    }

      $('#coffee-counter-container').show();
      $('#coffee-list-container').show();
      fadeCounter();
      fadeItem();

  //$('.ms-ListItem').on('click', clickFunc);
}

function fadeCounter()
{
    $('h3:hidden:first').fadeIn(1000,fadeItem);
}
function fadeItem() {
    $('ul li:hidden:first').fadeIn('fast',fadeItem);
}
