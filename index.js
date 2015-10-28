var contacts = new Array();

function log(texttolog) {
    var d = new Date();
    var time = padLeft(d.getHours(), 2) + ":" + padLeft(d.getMinutes(), 2) + ":" + padLeft(d.getSeconds(), 2) + ":" + padLeft(d.getMilliseconds(), 3);
    console.log(time + ": " + texttolog);
    $('#logging_box').html("<b>Status: </b>" + time + ": " + texttolog + "<br>");
}
function padLeft(nr, n, str) {
    return Array(n - String(nr).length + 1).join(str || '0') + nr;
}


function addToContacts(name, emailAddress, title, company, workPhone) {
    // Not every contact will have a title or company, 
    // so lets check for undefined values and display them more politely
    if (typeof (title) == "undefined") { title = 'n/a'; }
    if (typeof (company) == "undefined") { company = 'n/a'; }
    if (typeof (workPhone) == "undefined") { workPhone = 'n/a'; }
    // Comma delimited
    contacts.push('"' + name + '","' + emailAddress + '","' + title + '","' + company + '","' + workPhone + '"<br>');
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
    $('#btn').hide();

    log("App Loaded");
    $('#contacts').hide();

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

    log("Client Created");

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
            $('#contacts').show();            
           
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

    function retrieve_all() {
        log('Retrieving all contacts...');
        client.personsAndGroupsManager.all.persons.get().then(function (persons) {
            // `persons` is an array, so we can use Array::forEach here
            persons.forEach(function (person) {
                person.displayName.get().then(function (name) {
                    var personEmail = "";
                    person.emails.get().then(function (emails) {
                        var json_text = JSON.stringify(emails, null, 2).toString();
                        //log(name_id + ' : ' + json_text);
                        json_text = json_text.replace("[", "");
                        json_text = json_text.replace("]", "");
                        //log(name_id + ' : ' + json_text);
                        var obj = $.parseJSON(json_text);
                        //log(name_id + ' : ' + obj['emailAddress']);
                        var personEmail = obj['emailAddress'];
                        //add name_id and email address into array
                        addToContacts(name, personEmail, person.title(), person.company(), person.office());
                    });
                });
            });
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
                    $('#contacts').hide();
                },
            //onFailure callback
            function (error) {
                // or a failure
                log(error || 'Cannot Sign Out');
            });
    });

});