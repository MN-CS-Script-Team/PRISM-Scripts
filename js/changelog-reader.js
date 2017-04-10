function modifyDateRangeLast30Days() {
    // Writing values
    document.getElementById("start").value = moment().add(-30, 'days').format("MM/DD/YYYY");
    document.getElementById("end").value = moment().format("MM/DD/YYYY");
}

function modifyDateRangeLast90Days() {
    // Writing values
    document.getElementById("start").value = moment().add(-90, 'days').format("MM/DD/YYYY");
    document.getElementById("end").value = moment().format("MM/DD/YYYY");
}

function modifyDateRangeMonthToDate() {
    // Writing values
    document.getElementById("start").value = moment().format("MM/01/YYYY");
    document.getElementById("end").value = moment().format("MM/DD/YYYY");
}

function modifyDateRangeLastCompleteMonth() {
    // We need the first day of the current month, in order to determine the last day of the prior month, by subtracting 1 day
    var firstDayOfCurrentMonth = moment().format("MM/01/YYYY");

    // Writing values
    document.getElementById("start").value = moment().add(-1, 'months').format("MM/01/YYYY");
    document.getElementById("end").value = moment(firstDayOfCurrentMonth).add(-1, 'days').format("MM/DD/YYYY");
}

function msieversion() {

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE ");

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) { 
        var ieWarningHTML = document.getElementById("IEWarning");
        
        // This is the part that writes to the HTML doc
        ieWarningHTML.insertAdjacentHTML('beforeend', 
            // Here's the div I made up
            '<div class="alert alert-warning" role="alert"><span class="glyphicon glyphicon-warning-sign" aria-hidden="true"></span> You appear to be using Internet Explorer. This utility may not work in Internet Explorer. If nothing displays, please use Chrome or Firefox.</div>'
        );
    }

    return false;
}

function displaySpinner() {
    // Adds a loading spinner
    var goButtonColumn = document.getElementById("goButtonColumn");
    goButtonColumn.insertAdjacentHTML('beforeend', 
        //'<div id="loading"><img id="loading-image" src="img/loading.gif" alt="Loading..." /></div>'
        '<img id="loading-image" src="img/loading.gif" alt="Loading..." />'
    );
    
    // Now, run the changelog info
    displayChangelogInfo();
}

function displayChangelogInfo() {
        
    // Get the span for changelog contents, which is adds to later when we've retrieved details.
    var listOfScriptsHTML = document.getElementById("changelogContents");
    
    // Removes any existing details in the HTML doc (in case the report is re-run without refreshing)
    var tableRows = listOfScriptsHTML.getElementsByTagName('tr');
    var rowCount = tableRows.length;

    for (var x=rowCount; x>0; x--) {
        listOfScriptsHTML.removeChild(tableRows[x]);
    }
    
    
    

    
    // Storing the changelog_update string in a variable, which we'll use in our regex search later
    var functionToCheckFor = "changelog_update";
    
    // Gets the user-input from date picker, both begin and end
    var beginDateString = document.getElementById("start").value;
    var endDateString = document.getElementById("end").value;
    
    // Converts strings into proper date objects using moment.js
    var momentBeginDateObj = moment(beginDateString, 'MM/DD/YYYY');
    var momentEndDateObj = moment(endDateString, 'MM/DD/YYYY');
    
    // Then we need to know if the master branch will be used (it's a checkbox on the form)
    var masterBranchCheckbox = document.getElementById("scriptwriterBranchCheckbox");
    
    if (masterBranchCheckbox.checked) {
        var branchChoice = "master";
    } else {
        var branchChoice = "release";
    }
    
    // read text from URL location to get the list of scripts
    var request = new XMLHttpRequest();
    request.open('GET', 'https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/' + branchChoice + '/~complete-list-of-scripts.vbs', false);
    
    // This sends the request for info and does all of the hard work
    request.onreadystatechange = function () {
        // If the data is there, then...
        if (request.readyState === 4 && request.status === 200) {
            // create a new variable called "type" which handles the response header, or "type of content we're dealing with"
            var type = request.getResponseHeader('Content-Type');
            
            // If it's text, that means it's probably working and we can proceed!
            if (type.indexOf("text") !== 1) {
                
                // Create a variable filled with the contents of the FuncLib file
                var listOfScriptsArray = request.responseText.split("\n");
                
                var listOfScripts = "";
                
                
                for (var i = 0; i < listOfScriptsArray.length; i++) {
                    if (listOfScriptsArray[i].startsWith("cs_scripts_array(script_num).script_name")) {
                        
                        // Creating a friendly name for the new script
                        var scriptFriendlyName = listOfScriptsArray[i].replace('cs_scripts_array(script_num).script_name', '').replace(/"/g, '').replace('=', '').trim();
                        
                        // Getting the category, which is always on the next line
                        var scriptCategory = listOfScriptsArray[i + 1].slice((listOfScriptsArray[i + 1].length - listOfScriptsArray[i + 1].lastIndexOf("=")) * -1).replace(/"/g, '').replace('=', '').trim();
                        
                        // Getting the URL for the script file
                        var scriptURL = 'https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/'+ branchChoice + '/' + scriptCategory + '\\' + scriptFriendlyName.toLowerCase().replace(/ /g, '-') + '.vbs';
                                                
                        // read text from URL location to get the list of scripts
                        var scriptCheck = new XMLHttpRequest();
                            
                        // This sends the request for info and does all of the hard work
                        scriptCheck.onreadystatechange = function () {
                            // If the data is there, then...
                            if (scriptCheck.readyState === 4 && scriptCheck.status === 200) {
                                
                                // Gets the whole file
                                var data = scriptCheck.responseText;
                                
                                // Splits into array
                                var fileArray = data.split("\n");
                                
                                // This is a regex that checks for changelog_update
                                var re = new RegExp(functionToCheckFor, "i");

                                // Goes through the array created above, and checks for the changelog_update elements.
                                for (var j = 0; j < fileArray.length; j++) {
                                    if (fileArray[j].search(re) != -1) {
                                        
                                        // Regex for the changelog display, which is case insensitive
                                        var regexForChangelogDisplay = new RegExp("changelog_display", "i");
                                        
                                        // Escapes if the line meets certain criteria: 
                                        //      - it's an example
                                        //      - the string "changelog_display" is found (which means the end of the changelog block is here)
                                        if (fileArray[j] == "\'Example: call changelog_update(\"01/01/2000\", \"The script has been updated to fix a typo on the initial dialog.\", \"Jane Public, Oak County\")") {
                                            continue;
                                        } else if (fileArray[j].search(regexForChangelogDisplay) != -1) {
                                            break;
                                        }
                                        
                                        // Determines date, text, and scriptwriter
                                        var changelogEntryArray = fileArray[j].split('\"');         // splits into an array
                                        var changelogEntryDate = new Date(changelogEntryArray[1]);  // item [1] is the entry date
                                        var changelogEntryText = changelogEntryArray[3];            // item [3] is the entry text
                                        var changelogEntryScriptwriter = changelogEntryArray[5];    // item [5] is the entry scriptwriter
                                        
                                        // Uses moment.js to determine whether-or-not the script currently being evaluated falls within the date range specified by the user                                        
                                        var withinDateRange = moment(changelogEntryDate).isBetween(momentBeginDateObj, momentEndDateObj, 'day', '[]');
                                    
                                        // If we are within the range, it'll write to the HTML doc
                                        if (withinDateRange == true) {
                                            // This is the part that writes to the HTML doc
                                            listOfScriptsHTML.insertAdjacentHTML('beforeend', 
                                            
                                            "<tr> \n" +
                                            "    <td>" + scriptCategory.toUpperCase() + "</td> \n" + 
                                            "    <td><a href=\'" + scriptURL + "\' target=\'_blank\'>" + scriptFriendlyName + "</a></td> \n" + 
                                            //"    <td>" + changelogEntryDate.toDateString() + "</td> \n" + 
                                            "    <td>" + moment(changelogEntryDate).format('MM/DD/YYYY') + "</td> \n" + 
                                            "    <td>" + changelogEntryText + "</td> \n" + 
                                            "    <td>" + changelogEntryScriptwriter + "</td> \n" +
                                            "</tr>"
                                            );
                                        }
                                    };
                                }                                
                            } 
                        }
                        scriptCheck.open('GET', scriptURL, false);                            
                        scriptCheck.send();
                    }
                }
            }
        }
    }
    request.send(null);
    
    // Clear date selector
    var searchFormContents = document.getElementById("search-form");
    searchFormContents.innerHTML = "";
    
    var aboutContents = document.getElementById("about-this-list-text");
    aboutContents.innerHTML = "";
}

function sortTable(n) {
  var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
  table = document.getElementById("changelogTable");
  switching = true;
  //Set the sorting direction to ascending:
  dir = "asc"; 
  /*Make a loop that will continue until
  no switching has been done:*/
  while (switching) {
    //start by saying: no switching is done:
    switching = false;
    rows = table.getElementsByTagName("TR");
    /*Loop through all table rows (except the
    first, which contains table headers):*/
    for (i = 1; i < (rows.length - 1); i++) {
      //start by saying there should be no switching:
      shouldSwitch = false;
      /*Get the two elements you want to compare,
      one from current row and one from the next:*/
      x = rows[i].getElementsByTagName("TD")[n];
      y = rows[i + 1].getElementsByTagName("TD")[n];
      /*check if the two rows should switch place,
      based on the direction, asc or desc:*/
      if (dir == "asc") {
        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
          //if so, mark as a switch and break the loop:
          shouldSwitch= true;
          break;
        }
      } else if (dir == "desc") {
        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
          //if so, mark as a switch and break the loop:
          shouldSwitch= true;
          break;
        }
      }
    }
    if (shouldSwitch) {
      /*If a switch has been marked, make the switch
      and mark that a switch has been done:*/
      rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
      switching = true;
      //Each time a switch is done, increase this count by 1:
      switchcount ++; 
    } else {
      /*If no switching has been done AND the direction is "asc",
      set the direction to "desc" and run the while loop again.*/
      if (switchcount == 0 && dir == "asc") {
        dir = "desc";
        switching = true;
      }
    }
  }
}