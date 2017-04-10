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

function sortTable(n) {
    var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
    table = document.getElementById("scriptListTable");
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

$(document).ready(function(){

    // Setup config for key:value pairs needed
    var queryStringParams = { doc: [] };					// Query parameters
    var scriptListArray = { doc: [] };		// List of scripts in the collection
    
    // Setting up this variable, which is manipulated inside parseURLParams. This is used later to determine if fetchScript should be run.
    var queryStringFound = false;
    
    // Parse URL Params
    function parseURLParams(){
        $.each(location.search.replace(/\?/g,"").split(/\&/g), function(k, v) { 
            var param = v.split(/\=/g);
            queryStringParams[param[0]] = param[1]; 
            if (param[0] != "") {
                queryStringFound = true
            } else {
                queryStringFound = false
            };						// Sets an item for figuring out if a query string was found on the page
        });
    }
    
    // Sets up cross-site scripting, necessary for IE
    jQuery.support.cors = true;
    
    // Fetches the list of scripts
    function fetchListOfScripts(callback){
        $.ajax({
            url: "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/~complete-list-of-scripts.vbs",
            dataType: 'text',
            success: function( data ){
                callback(data);
            },
            error: function( data ){ 
                // Just a simple readout for now because we're very much testing this.
                alert(JSON.stringify(data));
            }
        });
    }

    function fetchScript(callback){
        $.ajax({
            url: "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/" + queryStringParams.script,
            dataType: 'text',
            success: function( data ){
                callback(data);
            },
            error: function( data ){ 
                // Just a simple readout for now because we're very much testing this.
                alert(JSON.stringify(data));
                
            }
        });
    }
    
    // Compile doc data into DOM
    function compileInstructionsOnPage(){
        // From original script
        $("#scriptListArea").remove();
        // Add elems to DOM
        $.each(queryStringParams.doc,function(k,v){
            if(v.tagName == "P") $('#testspan').append(v);
            else $('#testspan2').append(v);
        })	
    }
    
    // Compile doc data into DOM
    function compileScriptListOnPage(){
        
        // Add elems to DOM
        $.each(scriptListArray.doc,function(key,v){
            if(v.tagName == "TD") {
                if (v.className == "script-name") {
                    $('#scriptListContents').append("<tr>");
                } 
                
                $('#scriptListContents').append(v);
                
                if (v.className == "script-description") {
                    // <<<<<<<<<<<<<<<<<<< THE "MAYBE" TEXT IS TEMPORARY JUST TO DEMONSTRATE THE PROCESS
                    $('#scriptListContents').append("<td><a href='instructions.html?script=utilities%2Ffavorites-list.vbs'>Maybe</a></td></tr>");
                }
            } 			
        })	
    }
    
    // Parse URL queryStringParams 
    parseURLParams();
    
    // If a query string was passed through, run fetchscript, parse on success
    if (queryStringFound == true) {
        
        // Setup title
        $("#pageHeaderText").text("Instructions: " + decodeURIComponent(queryStringParams.script));
        
        fetchScript(function(script){
            //Parse Data
            queryStringParams.code = script.split(/\n/mg);
            $.each(queryStringParams.code,function(k,v){
                // Reviews the code for key elements and creates HTML elems based on those
                var elem, parsedData;
                if(v.indexOf("'~~~~~") > -1) {
                    parsedData = v.replace("'~~~~~","").trim();
                    elem = $('<p>',{ 'text': parsedData });
                } else if(v.indexOf("'+++++") > -1) {
                    parsedData = v.replace("'+++++","").trim();
                    elem = $('<img>',{ 'src': parsedData });
                } else if(v.indexOf("'*****") > -1) {
                    parsedData = v.replace("'*****","").trim();
                    elem = $('<figcaption>',{ 'text': parsedData, 'class': 'figure-caption' });
                }
                if(elem) queryStringParams.doc.push(elem[0]);
            });
            // Run DOM compiler
            compileInstructionsOnPage();
        });
    } else {
        fetchListOfScripts(function(script){
            //Parse Data
            scriptListArray.code = script.split(/\n/mg);
            $.each (scriptListArray.code, function (key, value) {
                // Reviews the code for key elements and creates HTML elems based on those
                var elem, parsedData;	// Clears and redefines
                if (value.indexOf("cs_scripts_array(script_num).script_name") > -1) {
                    parsedData = value.replace("cs_scripts_array(script_num).script_name","").replace("=", "").replace(/"/g , "").trim();
                    elem = $('<td>',{ 'text': parsedData, 'class': 'script-name'});
                } else if (value.indexOf("cs_scripts_array(script_num).category") > -1) {
                    parsedData = value.replace("cs_scripts_array(script_num).category","").replace("=", "").replace(/"/g , "").trim();
                    elem = $('<td>',{ 'text': parsedData, 'class': 'script-category' });
                } else if(value.indexOf("cs_scripts_array(script_num).description") > -1) {
                    parsedData = value.replace("cs_scripts_array(script_num).description","").replace("=", "").replace(/"/g , "").trim();
                    elem = $('<td>',{ 'text': parsedData, 'class': 'script-description' });
                }
                if (elem) scriptListArray.doc.push(elem[0]);
            });
            // Run DOM compiler
            compileScriptListOnPage();
        });
    }			
});