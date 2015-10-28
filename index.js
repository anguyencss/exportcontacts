// Array in which to store contacts
var contacts = new Array();

// Output activity log to the JavaScript console with a time/date stamp for debugging
function log(texttolog) {
    var d = new Date();
    var time = padLeft(d.getHours(), 2) + ":" + padLeft(d.getMinutes(), 2) + ":" + padLeft(d.getSeconds(), 2) + ":" + padLeft(d.getMilliseconds(), 3);
    console.log(time + ": " + texttolog);
    $('#logging_box').html("<b>Status: </b>" + time + ": " + texttolog + "<br>");
}
function padLeft(nr, n, str) {
    return Array(n - String(nr).length + 1).join(str || '0') + nr;
}

// Creates a comma delimited string for each contact found
function addToContacts(name, emailAddress, title, company, workPhone) {
    // Not every contact will have a title or company, 
    // so lets check for undefined values and display them more politely
    if (typeof (title) == "undefined") { title = 'n/a'; }
    if (typeof (company) == "undefined") { company = 'n/a'; }
    if (typeof (workPhone) == "undefined") { workPhone = 'n/a'; }
    // Comma delimited
    contacts.push('"' + name + '","' + emailAddress + '","' + title + '","' + company + '","' + workPhone + '"<br>');
    // Lets sort the contact alphabetically.
    contacts.sort();
    $('#export_box').html('name,email,title,company,telephone<br>' + contacts.join(""));
    log("");
    $('#logs').hide();
}


$(function () {
    'use strict';

    // new instance of clipboard.js
    var clipboard = new Clipboard('.btn');
    clipboard.on('success', function (e) {
        // clear selection after copying to clipboard
        e.clearSelection();
    });
    // Let's hide the Copy to Clipboard button until the export is finished.
    $('#btn').hide();

    log("App Loaded");

    var Application
    var client;
    Skype.initialize({
        apiKey: 'SWX-BUILD-SDK',
    }, function (api) {
        Application = api.application;
        client = new Application();
        
    }, function (err) {
        log('some error occurred: ' + err);
    });

    // Authenticates against a Lync or Skype for Business service
    function sign_in() {
        $('#signin').hide();
        log('Signing in...');
        // and invoke its asynchronous "signIn" method
        client.signInManager.signIn({
            username: $('#address').text(),
            password: $('#password').text()
        }).then(function () {
            log('Logged In Succesfully');
            $('#loginbox').hide();
        }).then(null, function (error) {
            // if either of the operations above fails, tell the user about the problem
            log(error || 'Oops, Something went wrong.');
            $('#signin').show()
        });
    }
    // when the user clicks the "Sign In" button
    $('#signin').click(function () {
        sign_in();
    });

    // Retrieves all contacts ('persons') asynchronously.
    // Note they do not return in any particular order, so they should be sorted
    function retrieve_all() {
        log('Retrieving all contacts...');
        client.personsAndGroupsManager.all.persons.get().then(function (persons) {
            // `persons` is an array, so we can use Array::forEach here
            persons.forEach(function (person) {
                person.displayName.get().then(function (name) {
                    var personEmail = "";
                    person.emails.get().then(function (emails) {
                        // a JSON string is returned containing one or more email addresses
                        var json_text = JSON.stringify(emails, null, 2).toString();
                        json_text = json_text.replace("[", "");
                        json_text = json_text.replace("]", "");
                        var obj = $.parseJSON(json_text);
                        var personEmail = obj['emailAddress'];
                        // Pass values to the addToContacts function that creates the CSV export
                        addToContacts(name, personEmail, person.title(), person.company(), person.office());
                    });
                });
            });
            // Once finished, we can show the copy button
            $('#btn').show();
        });
    }
    $('#retrieve_all').click(function () {
        retrieve_all();
    });

    // when the user clicks on the "Sign Out" button
    $('#signout').click(function () {
        // start signing out
        log("Signing Out");
        client.signInManager.signOut().then(
                //onSuccess callback
                function () {
                    // and report the success
                    log('Signed out');
                    $('#loginbox').show();
                    $('#signin').show();
                },
            //onFailure callback
            function (error) {
                // or a failure
                log(error || 'Cannot Sign Out');
            });
    });

});