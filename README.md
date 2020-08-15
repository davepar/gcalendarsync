# GCalendar Sync

A Google Sheet add-on for synchronizing events with Google Calendar.
[See the website](http://www.ballardsoftwarefoundry.com/gcalendarsync.html) for full instructions. The add-on is
in the [G Suite Marketplace](https://gsuite.google.com/marketplace/app/gcalendar_sync/831559814916).

## Contributing

The project started as a simple script to publish swim team practice times on a Google Calendar
and has evolved from there. Fixes and improvements are done by volunteers, so please be patient.

You're also welcome to dive into the code and send suggested changes and fixes as a pull request
or as text. To try your own changes, you'll need to install Clasp in order to compile the Typescript
and push it to a project. Follow the instructions
[here](https://developers.google.com/apps-script/guides/typescript).

Running the tests requires running a special pre/post script to modify imports:

    $ ./pretest
    $ npm test
    $ ./posttest

Find a bug? File an issue here on Github or email me via the link in the
[Contact section](http://www.ballardsoftwarefoundry.com/gcalendarsync.html#contact) of the website.
