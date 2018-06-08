(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

    loadCupsOfCoffee();

    };
  };

  function loadCupsOfCoffee()
  {
      $('#coffee-list-container').show();
      
  }
})();
