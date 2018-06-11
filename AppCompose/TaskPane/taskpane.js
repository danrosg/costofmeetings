var item;

(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    item = Office.context.mailbox.item;
    jQuery(document).ready(function(){

    //loadCupsOfCoffee();


      $('#insert-button').on('click', function(){

        getAllRecipients();
        buildCoffeeList("#coffee-list",10);

      })

    });
  };

})();


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
            write ('To-recipients of the item: ');
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
            write ('Cc-recipients of the item: ');
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
            write ('Bcc-recipients of the item: ');
            displayAddresses(asyncResult);
        }

        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
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
