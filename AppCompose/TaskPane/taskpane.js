var item;
var paygrades = {}; //hashmap for the paygrades
var userid_grades = {};
var recipients=[];
var total = 0;
var start = "";
var end = "";


(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    item = Office.context.mailbox.item;
    jQuery(document).ready(function(){
		// this is called once at document ready because for some reason the first time this function is called end time does not pull back. clicking the button will call this function a second time which pulls back start and end dates just fine.
		getHour();

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
      getCost();
      $('#coffee-list').empty();
      buildCoffeeList("#coffee-list",total);
      $('#coffee-fact-container').show();

		recipients=[];
    //total=0;

		//write(total);


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

    //var total=0;
	//var recipients=[];

    for (var i=0; i<asyncResult.value.length; i++)
    {

        var data =asyncResult.value[i].emailAddress;
        var userid = data.split('@');
        var name = userid.length==2 ? userid[0].toUpperCase() : null;

        //
        recipients.push(name);

        //total = total + userid_grades[name];
        //write ( name+' '+total+'\n');

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
    $('h3:hidden:first').fadeIn(4000,fadeCounter);
}
function fadeItem() {
    $('ul li:hidden:first').fadeIn(50,fadeItem);
}

// function to get start and end times of a meeting.
function getHour(){
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
			start = asyncResult.value;
			}
		});
	item.end.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
			end = asyncResult.value;
			}
		});
}

function getCost(){
  var storesalesperc;

	gradeTotal = 0;
	timeTotal = 0;
	total = 0;
	//Uses userid_grades to get grades of all users, then uses paygrades to get cost per hour.
	 for (var i=0; i<recipients.length; i++)
    {
		gradeTotal = gradeTotal + Number(paygrades[userid_grades[recipients[i]]]);

	}

	getHour();
	//write(gradeTotal);
    //calculates hours of meeting
	timeTotal = (Date.parse(end) - Date.parse(start))/1000/60/60;


	//write(gradeTotal);
	//sets total to be cost per hour gradeTotal times timeTotal
	total = ((gradeTotal * timeTotal)/3.25 ).toFixed(2);
  storesalesperc =(( total*3.25 )/4500*100 ).toFixed(2);

	document.getElementById('coffee-counter').innerText = total + ' Grande Americano cups';
  document.getElementById('coffee-fact-text').innerText = 'This represents '+storesalesperc+ '% of the average US Store Sales ....';
  document.getElementById('coffee-fact-text2').innerText = 'or '  +(storesalesperc*12/100).toFixed(2)+' hours of operation in a regular store';

}
