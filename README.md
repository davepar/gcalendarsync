# GCalendar Sync

**Update:** After 3 months of re-organizing and cleaning up the script, it's now in the process
of being approved for the GSuite Marketplace. Installing it from there will be much easier,
but isn't available yet (as of May 26, 2020). Until then you can install the older version of the
script located in the [priorversion](./priorversion) directory.

A Google Sheet Marketplace add-on for synchronizing events with Google Calendar. Instructions are located
on the [website](http://www.ballardsoftwarefoundry.com/gcalendarsync.html) for the add on.

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
