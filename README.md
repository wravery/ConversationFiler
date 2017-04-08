# ConversationFiler
Conversation Filer App for Outlook 2016

This is a rewrite of the [Conversation Filer App](https://store.office.com/en-us/app.aspx?assetid=WA103935787&ui=en-US&rs=en-US&ad=US&appredirect=false)
that I did for Outlook 2013 when the Office Store first launched. I originally wrote it in TypeScript 0.9 and used jQuery to implement the
[EWS](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html#makeEwsRequestAsync) SOAP requests and the UI, but I managed
to commit the cardinal sin of losing the source code in a freak laptop recycling accident. I was left with the transpiled JavaScript
version which I never updated. The App fell behind in terms of web technologies and Office APIs.

For this version, I reimplemented the UI with React.js, TypeScript 2.2, and I implemented a Data Access Layer which prefers the new [REST
APIs for Outlook](https://dev.outlook.com/). I translated the EWS Data module back to TypeScript and put it behind an identical facade,
because the Outlook 2016 Desktop application *[my day job, both versions of the App were hackathon projects]* doesn't expose the necessary
JS APIs to call the REST APIs. At runtime, if the App is hosted inside of a container that supports the REST APIs it will use them,
otherwise it will fall back to the ported EWS DAL from version 1.0.

The EWS API didn't have a good way to search the mailbox for conversations in the same folder, so it relied on reading the first 20
conversations in the Inbox. If the currently selected message was not in that first set, it would display an error. Besides being much
simpler to work with, the REST APIs let you search for all messages filtered down to the ones with a matching conversation ID, so there's
no more 20 conversation limit and the conversation doesn't need to be in the Inbox. If you are using the App in Outlook Web Access,
Outlook for Mac, or Outlook Mobile for iOS (as of today) it should never show you that particular error.

The latest Office JS APIs and manifest formats also enable a new preferred entrypoint for Apps: buttons which perform UI-less actions but
which can still show UI in a dialog. If you use v2.0 with Outlook 2013, it should still work as an inline App pane on individual messages,
but if you have Outlook 2016 or later it will use the new mechanism to integrate with the host. The new APIs also let me show progress
and notifications in the mail form, so going forward I'll be able to make the App show warnings and errors without even opening any new
UI surfaces.

The [manifest](./ConversationFiler.xml) file still needs some work, but if you do want to try running your own instance of the App you can
clone the repo, repoint all of the URLs in the manifest to your (HTTPS) endpoint, and then import the App from the manifest file in the
Manage Addins section of the Outlook Web Access settings pages. If you had v1.0 instsalled, this should overwrite it because it has the
same name and GUID identifier.
