require('dotenv').config();
const {google}=require('googleapis');
const express=require('express');
const bodyParser=require('body-parser');
var json2xls = require('json2xls');
const ejs=require('ejs');
const app=express();
const {OAuth2}=google.auth;


const oAuth2Client=new OAuth2(
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
)

oAuth2Client.setCredentials({
    refresh_token:process.env.REFRESH_TOKEN
})

const calendar=google.calendar({version:'v3', auth:oAuth2Client})


app.set("view engine", "ejs");
app.use(bodyParser.urlencoded({
  extended: true
}));
// app.use(express.static("public"));
app.use(express.static(__dirname+"/public"));

app.use(json2xls.middleware);

app.get("/",function(req,res){
    res.render("home");
})
app.get("/excel",function(req,res){
    res.render("excel");
})

app.get("/event",function(req,res){
    
    res.render("form");
})

app.post("/excel",function(req,result){
    const eventStartTime = new Date(req.body.meetingTimeFrom);
    const eventEndTime = new Date(req.body.meetingTimeTo);
    calendar.events.list({
        calendarId: 'primary',
        timeMin: eventStartTime,
        timeMax:eventEndTime,
        maxResults: 10,
        singleEvents: true,
        // orderBy: 'startTime',
        }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const events = res.data.items;
        console.log(events);
        result.xls('Events.xlsx',events,{
            fields: ['summary','description','location','creator.email','start.dateTime','end.dataTime']
        });
        // if (events.length) {
        //     console.log('Upcoming 10 events:');
        //     events.map((event, i) => {
        //       const start = event.start.dateTime || event.start.date;
        //       console.log(`${start} - ${event.summary}-${event.description}`);
        //     });
        //   } else {
        //     console.log('No upcoming events found.');
        //   }
    })
})
app.post("/event",function(req,result){
    // const eventStartTime = new Date("August 16, 2020 18:30:00")
    // // eventStartTime.setDate(eventStartTime.getDay()+2)
    
    // const eventEndTime=new Date("August 20, 2020 20:30:00")
    
    // eventEndTime.setDate(eventEndTime.getDay()+4);
    
    // eventEndTime.setMinutes(eventEndTime.getMinutes()+45)
    const eventStartTime = new Date(req.body.meetingTimeFrom);
    const eventEndTime = new Date(req.body.meetingTimeTo);
    const event = {
        summary: req.body.title,
        location: req.body.location,
        description: req.body.desc,
        colorId: Math.floor(Math.random()*11)+1,
        start: {
          dateTime: eventStartTime,
          timeZone: 'GMT+05:30',
        },
        end: {
          dateTime: eventEndTime,
          timeZone: 'GMT+05:30',
        },
        reminders: {
            useDefault: false,
            overrides: [
              {'method': 'popup', 
                'minutes': '2'},
            ],
          },
      }
      
      // Check if we a busy and have an event on our calendar for the same time.
      calendar.freebusy.query(
        {
          resource: {
            timeMin: eventStartTime,
            timeMax: eventEndTime,
            timeZone: 'GMT+05:30',
            items: [{ id: 'primary' }],
          },
        },
        (err, res) => {
          // Check for errors in our query and log them if they exist.
          if (err) return console.error('Free Busy Query Error: ', err)
      
          // Create an array of all events on our calendar during that time.
          const eventArr = res.data.calendars.primary.busy
      
          // Check if event array is empty which means we are not busy
          if (eventArr.length === 0)
            // If we are not busy create a new calendar event.
            return calendar.events.insert(
              { calendarId: 'primary', resource: event },
              err => {
                // Check for errors and log them if they exist.
                if (err) return console.error('Error Creating Calender Event:', err)
                // Else log that the event was created.
                return  result.render("eventCreated");
              }

            )
      
          // If event array is not empty log that we are busy.
          return console.log(`Sorry I'm busy...`)
        }
      )

    

})

app.listen(3000,function(){
    console.log("Server started at port 3000");
})