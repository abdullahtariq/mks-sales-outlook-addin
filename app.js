'use strict';


  Office.initialize = function (reason) {
    $(document).ready(function () {
      $("#run").html("init");
      var _mailbox = Office.context.mailbox;
       // Obtains the current item.
       $("#error").html("mail box");
       try {
       var item = _mailbox.item;
       var emailsHTML = "";
       var emailCount = 0;

       emailsHTML += "<li>"+item.sender.emailAddress+"</li>";
       emailCount = emailCount + 1;

       var toEmail = item.to;
       for (var i=0;i <toEmail.length;i++){
          emailsHTML += "<li>"+toEmail[i].emailAddress+"</li>";
          emailCount = emailCount + 1;
       }

       var ccEmail = item.cc;
       for (var i=0;i <ccEmail.length;i++){
          emailsHTML += "<li>"+ccEmail[i].emailAddress+"</li>";
          emailCount = emailCount + 1;
       }




       $(".emails ul").html(emailsHTML);
       $(".emailsfound").html(emailCount+" emails found in this message");

       $("#error").html(item.itemType);
       }
       catch(e){
         $("#error").html(e);
       }

    });
  };
