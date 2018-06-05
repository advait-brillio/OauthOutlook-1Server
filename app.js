
'use strict';
var restify = require('restify');
var builder = require('botbuilder');
var request = require('request-promise');
var apiai = require('apiai');
var app = apiai("4a667130013944bd988ed1a82959c1ab");
var authHelper = require('./authHelper');
var outlook = require('node-outlook');
const querystring = require('querystring');
var url = require('url');
var http = require('http');
var bodyParser = require('body-parser')


var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
console.log('%s is listening on port number ', server.name, server.url);
});


var connector = new builder.ChatConnector({
appId: "51f1c451-233b-4246-b40c-8d8e9315f8a1",
appPassword: "xzbqsIHV0?^mwQYKK1199=%"
});


server.post('/api/messages', connector.listen());
var mexp = require('math-expression-evaluator');

//weather modules and flags
var WeatherAPI = require('simple-weather-api');
var cityEntered = false;
var temp;
var min_temp;
var max_temp;
var humidity;
var w_speed;

//luis 
var bot = new builder.UniversalBot(connector);
var model = 'https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/b9ac6a7c-708b-41bd-a4b4-30d8855d4de3?subscription-key=576a1bf19110430b82391a00ffcdd655&verbose=true&timezoneOffset=0&q=';

var recognizer = new builder.LuisRecognizer(model);
var intents = new builder.IntentDialog({ recognizers: [recognizer] });


bot.dialog('/', intents);
bot.on('conversationUpdate', function (message) {
    if (message.membersAdded) {
        message.membersAdded.forEach(function (identity) {
            if (identity.id === message.address.bot.id) {
                bot.beginDialog(message.address, 'Introduce');
            }
        });
    }
});

bot.dialog('Introduce', [
    function (session, args, next) {
        console.log("session:!!!!!!!!", session.message.address);
        const commands = ' ### Hi Hella here ! I can help you with followings :\n - Type \"set language\" to speak in any languages(Currently Bot Support German,French,Spanish,Italian,Chinese,Japanese).\n - Type \"set default language\" to speak in english.\n - Type \" What is time? \"  to get time and date.\n - Type mathamatical questions and get result (Ex.add 1 and 2/5!/sin 45/log 10 etc)\n - Type  \"What is weather?\"  to see weather condition in city you required.\n - Type  \"Translate sentence\"  to translate any sentence to any other language.\n - Type  \"Translate sentence\"  to translate any sentence to any other language.\n - Type  \" get my mails \"  to get your mails.\n- Type  \" Introduce\"  at any time to see these options again.';
          session.send(commands).endDialog();
    }
]);
bot.dialog('Start Again', [
    function (session, args, next) {
        const commands = 'Language Changed, continue chatting..';
          session.send(commands).endDialog();
    }
]);




intents.matches('greetings', [function (session, args) {
    session.send('Hello');
    console.log("hiiiiiiiiiiii");
}]);



intents.matches('None', [function (session, args) {
    apiCall(session, args);
}])

var fs = require('fs');
var writtableStream=fs.createWriteStream('Suggestions.txt');

//================================================================================================================
var languageTo="";
function apiCall(session, args) {
            if(other_lang){
            translateQue(session,session.message.text);
              setTimeout(() => {
             var a = que;
             console.log("other "+a)
                        var request = app.textRequest(a, {
                                sessionId: '1234567891'
                        });



        request.on('response', function (response) {
                var a = response;

                if (a["result"]["metadata"]["intentName"] == "smalltalk.math") {
                        solveMath(session, args, a);
                        autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                }  
                else if (a["result"]["metadata"]["intentName"] == "Time") {
                        let date = require('date-and-time');
                        let now = new Date();
                        // date.format(now, 'YYYY/MM/DD HH:mm:ss');    // => '2015/01/02 23:14:05'
                        // date.format(now, 'ddd MMM DD YYYY');        // => 'Fri Jan 02 2015'
                        // date.format(now, 'hh:mm A [GMT]Z');         // => '11:14 p.m. GMT-0800'
                        // date.format(now, 'hh:mm A [GMT]Z', true);   // => '07:14 a.m. GMT+0000'
                        var pre=""
                        if(date.format(now, "D")=="1"||date.format(now, "D")=="21"||date.format(now, "D")=="31")pre="st"
                        else if(date.format(now, "D")=="2"||date.format(now, "D")=="22")pre="nd"
                        else if(date.format(now, "D")=="3"||date.format(now, "D")=="23")pre="rd"
                        else pre="th"
                        session.send("It is "+date.format(now, "h:mm A")+" now" );
                        autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                }
                else if (a["result"]["metadata"]["intentName"] == "smalltalk.date") {
                       var day = a["result"]["parameters"]["date"]
                        let date = require('date-and-time');
                        let now = new Date();

                        var pre=""
                        if(date.format(now, "D")=="1"||date.format(now, "D")=="21"||date.format(now, "D")=="31")pre="st"
                        else if(date.format(now, "D")=="2"||date.format(now, "D")=="22")pre="nd"
                        else if(date.format(now, "D")=="3"||date.format(now, "D")=="23")pre="rd"
                        else pre="th";

                        if(day=="tomorrow"){
                        var yesterday = date.addDays(now, +1);
                        session.send("Tomorrow is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"));
                        }
                        else if(day=="yesterday"){
                        var yesterday = date.addDays(now, -1);
                        session.send("Yesterday is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                        }
                        else if(day=="day after yesterday"){
                        var yesterday = date.addDays(now, 0);
                        session.send("Day after yesterday is Today, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                        }
                        else if(day=="day before today"){
                        var yesterday = date.addDays(now, -1);
                        session.send("Day before today is yesterday, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day after tomorrow"){
                        var yesterday = date.addDays(now, +2);
                        session.send("Day after tomorrow is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day before tomorrow"){
                        var yesterday = date.addDays(now, 0);
                        session.send("Day before tomorrow is today, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day after today"){
                        var yesterday = date.addDays(now, +1);
                        session.send("Day after today is tomorrow, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day before yesterday"){
                        var yesterday = date.addDays(now, -2);
                        session.send("Day before yesterday is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else
                        session.send("Today is "+date.format(now, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                }
                else if (a["result"]["metadata"]["intentName"] == "Weather") {
                        var city = a["result"]["parameters"]["geo-city"]
                        console.log(a);
                        if (city) {
                            var apikey = "233becf468d81bc10b7f315b126c8846";
                            var weather = new WeatherAPI(apikey);
                            weather.getWeather(city).then(response => {
                                console.log("response....." + response.body);
                                let san = JSON.parse(response.body);
                                temp = (san.main.temp - 273.15).toFixed(2);
                                min_temp = (san.main.temp_min - 273.15).toFixed(2);
                                max_temp = (san.main.temp_max - 273.15).toFixed(2);
                                humidity = (san.main.humidity);
                                w_speed = (san.wind.speed);
                                session.send("It's " + temp + " degrees celsius in " + city);
                                weatherCity(session, args);
                            });
                        } else {
                                cityEntered = true;
                                session.send('Which city do you want the weather for?');
                        }
                } 
                else if (cityEntered) {
                        console.log("city entered");
                        cityEntered = false;
                        city = a["result"]["resolvedQuery"];
                        var apikey = "233becf468d81bc10b7f315b126c8846";
                        var weather = new WeatherAPI(apikey);
                        weather.getWeather(city).then(response => {
                                console.log("response....." + response.body);
                                let san = JSON.parse(response.body);
                                temp = (san.main.temp - 273.15).toFixed(2);
                                min_temp = (san.main.temp_min - 273.15).toFixed(2);
                                max_temp = (san.main.temp_max - 273.15).toFixed(2);
                                humidity = (san.main.humidity);
                                w_speed = (san.wind.speed);
                                session.send("It's " + temp + " degrees celsius in " + city);
                                weatherCity(session, args);
                                autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                        });
                }
                else if (a["result"]["metadata"]["intentName"] == "hella_user_outlook") {
                    session.beginDialog("get my mails")
                }
                else if (a["result"]["metadata"]["intentName"] == "smalltalk.translate") {
                    Translate_lang_flag=true;
                    var lang = a["result"]["parameters"]["language"]
                    if(lang=="German")languageTo="de"
                    else if(lang=="French")languageTo="fr"
                    else if(lang=="Spanish")languageTo="es"
                    else if(lang=="Italian")languageTo="it"
                    else if(lang=="Japanese")languageTo="ja"
                    else if(lang=="Chinese")languageTo="zh-Hans"
                    else languageTo=""
                    
                    session.beginDialog("Translate sentence")
                }
                        else {


                            // var p="";
                            //                 if(lang_other=="German")p="de"
                            //                 else if(lang_other=="French")p="fr"
                            //                 else if(lang_other=="Spanish")p="es"
                            //                 else if(lang_other=="Italian")p="it"
                            //                 else if(lang_other=="Japanese")p="ja"
                            //                 else if(lang_other=="Chinese")p="zh-Hans"
                            //                 else p="en";
                            //                 setTimeout(() => {
                            //                         translateAns(session,a["result"]["fulfillment"]["speech"],p);
                            //                         // session.send(a["result"]["fulfillment"]["speech"])
                            //                 }, 1000);
                    
                            //         }
                        var intent_name=a["result"]["metadata"]["intentName"]
                        var arr=[];
                        arr=intent_name.split(".");
                        if(arr.length<=1)intent_name1=arr[0];
                        var intent_name1=arr[0]+"."+arr[1];
                        session.sendTyping();
                         setTimeout(function () {
                                 autoSuggest(session,intent_name1,"3e303086904f452facfea90b6e014bab",a["result"]["fulfillment"]["speech"]);
                            }, 3000);
                       }
        });

        request.on('error', function (error) {
        console.log(error);
        });

        request.end();
        return;
        }, 1000);
}
else{
            var a = session.message.text;
            console.log("normal "+a)
        var request = app.textRequest(a, {
                sessionId: '1234567891'
        });

          request.on('response', function (response) {
                var a = response;

                if (a["result"]["metadata"]["intentName"] == "smalltalk.math") {
                        solveMath(session, args, a);
                        autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                }  
                else if (a["result"]["metadata"]["intentName"] == "Time") {
                        let date = require('date-and-time');
                        let now = new Date();
                        // date.format(now, 'YYYY/MM/DD HH:mm:ss');    // => '2015/01/02 23:14:05'
                        // date.format(now, 'ddd MMM DD YYYY');        // => 'Fri Jan 02 2015'
                        // date.format(now, 'hh:mm A [GMT]Z');         // => '11:14 p.m. GMT-0800'
                        // date.format(now, 'hh:mm A [GMT]Z', true);   // => '07:14 a.m. GMT+0000'
                        var pre=""
                        if(date.format(now, "D")=="1"||date.format(now, "D")=="21"||date.format(now, "D")=="31")pre="st"
                        else if(date.format(now, "D")=="2"||date.format(now, "D")=="22")pre="nd"
                        else if(date.format(now, "D")=="3"||date.format(now, "D")=="23")pre="rd"
                        else pre="th";
                        session.send("It is "+date.format(now, "h:mm A")+" now" )
                        // token	meaning	example
                        // YYYY	year	0999, 2015
                        // YY	year	15, 99
                        // Y	year	999, 2015
                        // MMMM	month	January, December
                        // MMM	month	Jan, Dec
                        // MM	month	01, 12
                        // M	month	1, 12
                        // DD	day	02, 31
                        // D	day	2, 31
                        // dddd	day of week	Friday, Sunday
                        // ddd	day of week	Fri, Sun
                        // dd	day of week	Fr, Su
                        // HH	hour-24	23, 08
                        // H	hour-24	23, 8
                        // A	meridiem	a.m., p.m.
                        // hh	hour-12	11, 08
                        // h	hour-12	11, 8
                        // mm	minute	14, 07
                        // m	minute	14, 7
                        // ss	second	05, 10
                        // s	second	5, 10
                        // SSS	millisecond	753, 022
                        // SS	millisecond	75, 02
                        // S	millisecond	7, 0
                        // Z	timezone	+0100, -0800
                        autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                }
                  else if (a["result"]["metadata"]["intentName"] == "smalltalk.date") {
                    var day = a["result"]["parameters"]["date"]
                        let date = require('date-and-time');
                        let now = new Date();

                        var pre=""
                        if(date.format(now, "D")=="1"||date.format(now, "D")=="21"||date.format(now, "D")=="31")pre="st"
                        else if(date.format(now, "D")=="2"||date.format(now, "D")=="22")pre="nd"
                        else if(date.format(now, "D")=="3"||date.format(now, "D")=="23")pre="rd"
                        else pre="th";

                        if(day=="tomorrow"){
                        var yesterday = date.addDays(now, +1);
                        session.send("Tomorrow is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"));
                        }
                        else if(day=="yesterday"){
                        var yesterday = date.addDays(now, -1);
                        session.send("Yesterday is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                        }
                        else if(day=="day after yesterday"){
                        var yesterday = date.addDays(now, 0);
                        session.send("Day after yesterday is Today, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                        }
                        else if(day=="day before today"){
                        var yesterday = date.addDays(now, -1);
                        session.send("Day before today is yesterday, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day after tomorrow"){
                        var yesterday = date.addDays(now, +2);
                        session.send("Day after tomorrow is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day before tomorrow"){
                        var yesterday = date.addDays(now, 0);
                        session.send("Day before tomorrow is today, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day after today"){
                        var yesterday = date.addDays(now, +1);
                        session.send("Day after today is tomorrow, that is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else if(day=="day before yesterday"){
                        var yesterday = date.addDays(now, -2);
                        session.send("Day before yesterday is "+date.format(yesterday, "dddd  D")+pre+date.format(now, " MMMM, Y"))
                        }
                        else
                        session.send("Today is "+date.format(now, "dddd  D")+pre+date.format(now, " MMMM, Y")) 
                }
                else if (a["result"]["metadata"]["intentName"] == "Weather") {
                        var city = a["result"]["parameters"]["geo-city"]
                        console.log(a);
                        if (city) {
                            var apikey = "233becf468d81bc10b7f315b126c8846";
                            var weather = new WeatherAPI(apikey);
                            weather.getWeather(city).then(response => {
                                console.log("response....." + response.body);
                                let san = JSON.parse(response.body);
                                temp = (san.main.temp - 273.15).toFixed(2);
                                min_temp = (san.main.temp_min - 273.15).toFixed(2);
                                max_temp = (san.main.temp_max - 273.15).toFixed(2);
                                humidity = (san.main.humidity);
                                w_speed = (san.wind.speed);
                                session.send("It's " + temp + " degrees celsius in " + city);
                                weatherCity(session, args);
                            });
                        } else {
                                cityEntered = true;
                                session.send('Which city do you want the weather for?');
                        }
                } 
                else if (cityEntered) {
                        console.log("city entered");
                        cityEntered = false;
                        city = a["result"]["resolvedQuery"];
                        var apikey = "233becf468d81bc10b7f315b126c8846";
                        var weather = new WeatherAPI(apikey);
                        weather.getWeather(city).then(response => {
                                console.log("response....." + response.body);
                                let san = JSON.parse(response.body);
                                temp = (san.main.temp - 273.15).toFixed(2);
                                min_temp = (san.main.temp_min - 273.15).toFixed(2);
                                max_temp = (san.main.temp_max - 273.15).toFixed(2);
                                humidity = (san.main.humidity);
                                w_speed = (san.wind.speed);
                                session.send("It's " + temp + " degrees celsius in " + city);
                                weatherCity(session, args);
                                autoSuggest(session,"","4a667130013944bd988ed1a82959c1ab",a["result"]["fulfillment"]["speech"]);
                        });
                }
                else if (a["result"]["metadata"]["intentName"] == "hella_user_outlook") {
                    session.beginDialog("get my mails")
                }
                else if (a["result"]["metadata"]["intentName"] == "smalltalk.translate") {
                    Translate_lang_flag=true;
                    var lang = a["result"]["parameters"]["language"]
                    if(lang=="German")languageTo="de"
                    else if(lang=="French")languageTo="fr"
                    else if(lang=="Spanish")languageTo="es"
                    else if(lang=="Italian")languageTo="it"
                    else if(lang=="Japanese")languageTo="ja"
                    else if(lang=="Chinese")languageTo="zh-Hans"
                    else languageTo=""
                    
                    session.beginDialog("Translate sentence")
                }

               else {
                //    var p="";
                //    if(lang_other=="German")p="de"
                //     else if(lang_other=="French")p="fr"
                //     else if(lang_other=="Spanish")p="es"
                //     else if(lang_other=="Italian")p="it"
                //     else if(lang_other=="Japanese")p="ja"
                //     else if(lang_other=="Chinese")p="zh-Hans"
                //     else p="en";
                //     setTimeout(() => {
                //             translateAns(session,a["result"]["fulfillment"]["speech"],p);
                //             // session.send(a["result"]["fulfillment"]["speech"])
                //     }, 1000);

                    var intent_name=a["result"]["metadata"]["intentName"]
                        var arr=[];
                        arr=intent_name.split(".");
                        if(arr.length<=1)intent_name1=arr[0];
                        var intent_name1=arr[0]+"."+arr[1];
                        session.sendTyping();
                         setTimeout(function () {
                                 autoSuggest(session,intent_name1,"3e303086904f452facfea90b6e014bab",a["result"]["fulfillment"]["speech"]);
                            }, 3000);
                
    
                    }
        });

        request.on('error', function (error) {
        console.log(error);
        });

        request.end();
        return;
    }
}
//=====================================

//weather suggestion function

function weatherCity(session, args) {
    session.sendTyping();
    var msg = new builder.Message(session)
    .text("What else do you want to see?")
    .suggestedActions(
    builder.SuggestedActions.create(
        session, [
        builder.CardAction.imBack(session, "I want to see minimum temperature", "Minimum Temperature"),
        builder.CardAction.imBack(session, "I want to see maximum temperature", "Maximum Temperature"),
        builder.CardAction.imBack(session, "I want to see humidity", "Humidity"),
        builder.CardAction.imBack(session, "I want to see wind speed", "Wind Speed")
        ]
    ));
    session.send(msg);
}



bot.dialog('minimum temperature', [
function (session, args) {
    session.sendTyping();
    setTimeout(() => {
    session.send("Minimum temperature is " + min_temp + " degrees celsius").endDialog();
    }, 1000);
}
]).triggerAction({ matches: /I want to see minimum temperature/i });


bot.dialog('maximum temperature', [

function (session, args) {
    session.sendTyping();
    setTimeout(() => {
    session.send("Maximum temperature is " + max_temp + " degrees celsius").endDialog();
    }, 1000);
}
]).triggerAction({ matches: /I want to see maximum temperature/i });



bot.dialog('humidity', [
function (session, args) {
    session.sendTyping();
    setTimeout(() => {
    session.send("Humidity is " + humidity + "%").endDialog();
    }, 1000);
}
]).triggerAction({ matches: /I want to see humidity/i });


bot.dialog('wind speed', [
function (session, args) {
    session.sendTyping();
    setTimeout(() => {
    session.send("Wind speed is " + w_speed + "km/h").endDialog();
    }, 1000);
}
]).triggerAction({ matches: /I want to see wind speed/i });


//math
function solveMath(session, args, a) {
var sum = 0;
var expression = "", expression1 = "", expression2 = "";
var i = 0, j = 0, k = 0;
expression = a["result"]["resolvedQuery"];

var numbers = [{ "name": "one", "value": "1" }, { "name": "two", "value": "2" }, { "name": "three", "value": "3" }, { "name": "four", "value": "4" }, { "name": "five", "value": "5" }, { "name": "six", "value": "6" }, { "name": "seven", "value": "7" }, { "name": "eight", "value": "8" }, { "name": "nine", "value": "9" }, { "name": "ten", "value": "10" }, { "name": "zero", "value": "0" }, { "name": "plus", "value": "+" }, { "name": "minus", "value": "-" }];

for (var s = 0; s < numbers.length; s++) {
    while (expression.indexOf(numbers[s].name) > -1)
    expression = expression.replace(numbers[s].name, numbers[s].value);
}

if (expression.indexOf("sum of") > -1||expression.indexOf("add") > -1||expression.indexOf("Add") > -1||expression.indexOf("Sum of") > -1) {
            var q = 0, sum = 0;
            expression1 = "Sum of "
            var numberPattern = /\d+/g;
            var nums = expression.match(numberPattern);
            while (nums[q] != null) {

                sum += parseInt(nums[q])
                if (expression1 == "Sum of ")
                expression1 += " " + nums[q] + " ";
                else
                expression1 += "+" + nums[q] + " "
                q++;

             }
        }
else if (expression.indexOf("Subtract") > -1||expression.indexOf("subtract") > -1) {
            var q = 0, sum = 0;
            expression1 = " Result "
            var numberPattern = /\d+/g;
            var nums = expression.match(numberPattern);
            while (nums[q] != null) {
                if(q==0)
                    sum = parseInt(nums[q])
                else
                    sum -= parseInt(nums[q])
                q++;
            }
        }
else if (expression.indexOf("product of") > -1) {
        var q = 0, sum = 1;
        expression1 = "Product of "
        var numberPattern = /\d+/g;
        var nums = expression.match(numberPattern);
        while (nums[q] != null) {
            sum *= parseInt(nums[q])
            expression1 += " " + nums[q] + ","
            q++;
        }
        }
else {
        if (expression.indexOf("what is") > -1) {
                i = expression.indexOf("what is") + "what is".length;
        }

        if (expression.indexOf("solve") > -1) {
             i = expression.indexOf("solve") + "solve".length;
        }
        while (expression[i] != null) {
            expression1 += expression[i];
            i++;
        }
        var sum = mexp.eval(expression1);
}

        session.send(expression1 + ' is ' + sum);
        return;
}







function autoSuggest(session,intent_name,key,resp) {
request({
        "method":"GET", 
        "uri": "https://api.dialogflow.com/v1/intents?v=20150910",
        "json": true,
        "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer "+key
}
}).then(function(response) {
        var i=0,j=0,matched_intents_id=[],Weather="",math="";
        var pattern="",match=false;
        while(response[i].name!=null){
                if(response[i].name=="Weather")Weather=response[i].id;
                if(response[i].name=="smalltalk.math")math=response[i].id;
                pattern = new RegExp(intent_name,"i");
                match=pattern.test(response[i].name);
                if(match){
                    matched_intents_id[j]=response[i].id;
                    j++;
                }
                if(i>=response.length-1)
                    break;
                else
                    i++;
        }
        var lower_bound = 0,
        upper_bound = matched_intents_id.length;
        if(upper_bound<=1)
            session.send(resp)
        else
            get_questions(session,matched_intents_id,key,lower_bound,upper_bound,resp,Weather,math)
})
}



function get_questions(session,matched_intents_id,key,lower_bound,upper_bound,resp,Weather,math){
        var random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
        request({
                "method":"GET", 
                "uri": "https://api.dialogflow.com/v1/intents/"+ matched_intents_id[random_number],
                "json": true,
                "headers": {
                "Content-Type": "application/json",
                "Authorization": "Bearer "+key
        }
        }).then(function(response) {
                var data=[];
                var j=0;
                while(response.userSays[j]!=null){
                    var k=0,text="";
                    while(response.userSays[j].data[k]!=null){
                        text+=" "+response.userSays[j].data[k].text;
                        k++;
                    }
                    data[j]=text;
                    j++;
                }
                return data
        }).then(function(response) {
                random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
                get_questions2(session,matched_intents_id,key,lower_bound,upper_bound,response[0],resp,Weather,math)
        }); 
}

function get_questions2(session,matched_intents_id,key,lower_bound,upper_bound,response0,resp,Weather,math){
        var random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
        request({
                "method":"GET", 
                "uri": "https://api.dialogflow.com/v1/intents/"+ matched_intents_id[random_number],
                "json": true,
                "headers": {
                "Content-Type": "application/json",
                "Authorization": "Bearer "+key
        }
        }).then(function(response) {
                var data=[];
                var j=0;
                while(response.userSays[j]!=null){
                    var k=0,text="";
                    while(response.userSays[j].data[k]!=null){
                        text+=" "+response.userSays[j].data[k].text;
                        k++;
                    }
                    data[j]=text;
                    j++;
                }
                return data
        }).then(function(response) {
                random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
                get_questions3(session,matched_intents_id,key,lower_bound,upper_bound,response0,response[0],resp,Weather,math)
}); 
}


function get_questions3(session,matched_intents_id,key,lower_bound,upper_bound,response0,response1,resp,Weather,math){
        var random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
        request({
                "method":"GET", 
                "uri": "https://api.dialogflow.com/v1/intents/"+ matched_intents_id[random_number],
                "json": true,
                "headers": {
                "Content-Type": "application/json",
                "Authorization": "Bearer "+key
        }
        }).then(function(response) {
                var data=[];
                var j=0;
                while(response.userSays[j]!=null){
                        var k=0,text="";
                        while(response.userSays[j].data[k]!=null){
                            text+=" "+response.userSays[j].data[k].text;
                            k++;
                        }
                        data[j]=text;
                        j++;
                }
                return data
        }).then(function(response) {
                random_number = Math.floor(Math.random()*(upper_bound - lower_bound) + lower_bound);
                get_questions4(session,matched_intents_id,key,lower_bound,upper_bound,response0,response1,response[0],resp,Weather,math
                )
        }); 
}


function get_questions4(session,matched_intents_id,key,lower_bound,upper_bound,response0,response1,response2,resp,Weather,math){
        request({
                "method":"GET", 
                "uri": "https://api.dialogflow.com/v1/intents/"+ Weather,
                "json": true,
                "headers": {
                "Content-Type": "application/json",
                "Authorization": "Bearer "+key
        }
        }).then(function(response) {
                var data=[];
                var j=0;
                while(response.userSays[j]!=null){
                    var k=0,text="";
                    while(response.userSays[j].data[k]!=null){
                        text+=" "+response.userSays[j].data[k].text;
                        k++;
                    }
                    data[j]=text;
                    j++;
                }
                return data
        }).then(function(response) {
                var random_number = Math.floor(Math.random()*(response.length - lower_bound) + lower_bound);
                get_questions5(session,matched_intents_id,key,lower_bound,upper_bound,response0,response1,response2,response[random_number],resp,Weather,math)
        }); 
}

function get_questions5(session,matched_intents_id,key,lower_bound,upper_bound,response0,response1,response2,response3,resp,Weather,math){
        request({
                "method":"GET", 
                "uri": "https://api.dialogflow.com/v1/intents/"+ math,
                "json": true,
                "headers": {
                "Content-Type": "application/json",
                "Authorization": "Bearer "+key
        }
        }).then(function(response) {
                var data=[];
                var j=0;
                while(response.userSays[j]!=null){
                    var k=0,text="";
                    while(response.userSays[j].data[k]!=null){
                        text+=" "+response.userSays[j].data[k].text;
                        k++;
                    }
                    data[j]=text;
                    j++;
                }
                return data
        }).then(function(response) {
                var random_number = Math.floor(Math.random()*(response.length - lower_bound) + lower_bound);
                display(session,response0,response1,response2,response3,response[random_number],resp)
                }); 
}

function display(session,response0,response1,response2,response3,response4,resp){
    if(other_lang){
        // response0=response0.toUpperCase();
        // response1=response1.toUpperCase();
        // response2=response2.toUpperCase();
        // response3=response3.toUpperCase();
        var p="";
            if(lang_other=="German")p="de"
            else if(lang_other=="French")p="fr"
            else if(lang_other=="Spanish")p="es"
            else if(lang_other=="Italian")p="it"
            else if(lang_other=="Japanese")p="ja"
            else if(lang_other=="Chinese")p="zh-Hans"
            else p="en";
            var x=response0.toUpperCase()+" 123 "+response1.toUpperCase()+" 123 "+response2.toUpperCase()+" 123 "+response3.toUpperCase()+" 123 "+response4.toUpperCase()+" 123 "+resp
           console.log("other "+x)
            translateSug(session,x,p);
              setTimeout(() => {
             var a = que;
             console.log("other "+a)
               var s_arr=[];
               s_arr=a.split("123")         
        var msg = new builder.Message(session)
        .text(s_arr[5])
        .suggestedActions(
        builder.SuggestedActions.create(
            session, [
            builder.CardAction.imBack(session, s_arr[0], s_arr[0]),
            builder.CardAction.imBack(session, s_arr[1], s_arr[1]),
            builder.CardAction.imBack(session, s_arr[2], s_arr[2]),
            builder.CardAction.imBack(session, s_arr[3], s_arr[3]),
            builder.CardAction.imBack(session, response4, response4),
            builder.CardAction.imBack(session, "Translate sentence", "Translate sentence"),
            ]
        ));
        
        session.send(msg);
        }, 1000);
    }
    else
    {
          response0=response0.toUpperCase();
        response1=response1.toUpperCase();
        response2=response2.toUpperCase();
        response3=response3.toUpperCase();
        var msg = new builder.Message(session)
        .text(resp)
        .suggestedActions(
        builder.SuggestedActions.create(
            session, [
            builder.CardAction.imBack(session, response0, response0),
            builder.CardAction.imBack(session, response1, response1),
            builder.CardAction.imBack(session, response2, response2),
            builder.CardAction.imBack(session, response3, response3),
            builder.CardAction.imBack(session, response4, response4),
            builder.CardAction.imBack(session, "Translate sentence", "Translate sentence"),
            ]
        ));
        
        session.send(msg);
    }
}
var message="",Translate_flag=false,Translate_lang_flag=false;
bot.dialog('Translate sentence', [
     function (session) {
         if(Translate_flag==0){session.send("Type 'Translate sentence' at any time to translation");Translate_flag=true;}
        builder.Prompts.text(session, 'What is the sentence that you want to translate?');
    },
    function (session, results,next) {
        message=results.response;
        next();
    },
    function (session,next) {
        session.sendTyping();
        setTimeout(() => {
            if(Translate_lang_flag){
                if(languageTo!="")
                translateMe(session,message,languageTo);
                else
                langtranslate(session)
                setTimeout(() => {
           Translate_lang_flag=false;
        }, 10);}
            else
            langtranslate(session)
        }, 1000);
    }
]).triggerAction({ matches: /Translate sentence/i });

function langtranslate(session) {
    session.sendTyping();
    var msg = new builder.Message(session)
        .text("Which language do you want to translate into?")
        .suggestedActions(
        builder.SuggestedActions.create(
            session, [
                builder.CardAction.imBack(session, "I want to translate into German", "German"),
                builder.CardAction.imBack(session, "I want to translate into French", "French"),
                builder.CardAction.imBack(session, "I want to translate into Spanish", "Spanish"),
                builder.CardAction.imBack(session, "I want to translate into Italian", "Italian"),
                builder.CardAction.imBack(session, "I want to translate into Japanese", "Japanese"),
                builder.CardAction.imBack(session, "I want to translate into Chinese", "Chinese"),
            ]
        ));
    session.send(msg);
}


bot.dialog('German', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"de");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into German/i });



bot.dialog('French', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"fr");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into French/i });




bot.dialog('Spanish', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"es");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into Spanish/i });



bot.dialog('Italian', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"it");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into Italian/i });

bot.dialog('Japanese', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"ja");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into Japanese/i });

bot.dialog('Chinese', [
    function (session) {
        session.sendTyping();
        setTimeout(() => {
            translateMe(session,message,"zh-Hans");
        }, 1000);
    }
]).triggerAction({ matches: /I want to translate into Chinese/i });










function translateMe(session,textToTranslate,p) {
console.log(textToTranslate)
//langtranslate(session,session.message.text);
let fs = require ('fs');
let https = require ('https');
let subscriptionKey = 'b7bfbd192aa5442eb355e828444ace35';
let host = 'api.cognitive.microsofttranslator.com';
let path = '/translate?api-version=3.0';
// Translate to German and Italian.
//let params = '&to=de&to=it&to=es&to=fr&to=ja&to=zh';
let params = '&to='+p;
let text = textToTranslate;
let response_handler = function (response) {
    let body = '';
    response.on ('data', function (d) {
        body += d;
    });
    response.on ('end', function () {
        let json = JSON.stringify(JSON.parse(body), null, 4);
        console.log(json);
        var i=0;
        while(JSON.parse(body)[i]!=null)
        {
            var j=0;
            while(JSON.parse(body)[i].translations[j]!=null){
                if(JSON.parse(body)[i].translations[j].to=="de")
                var lang="German"
                else if(JSON.parse(body)[i].translations[j].to=="it")
                var lang="Italian"
                else if(JSON.parse(body)[i].translations[j].to=="es")
                var lang="Spanish"
                else if(JSON.parse(body)[i].translations[j].to=="fr")
                var lang="French"
                 else if(JSON.parse(body)[i].translations[j].to=="ja")
                var lang="Japanese"
                 else if(JSON.parse(body)[i].translations[j].to=="zh-Hans")
                var lang="Chinese"
                message="";
                session.send(JSON.parse(body)[i].translations[j].text ).endDialog();;
                
                // session.beginDialog('begin');
                j++;
            }
            i++;

        }
    });
    response.on ('error', function (e) {
        console.log ('Error: ' + e.message);
    });
};

let get_guid = function () {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

let Translate = function (content) {
    let request_params = {
        method : 'POST',
        hostname : host,
        path : path + params,
        headers : {
            'Content-Type' : 'application/json',
            'Ocp-Apim-Subscription-Key' : subscriptionKey,
            'X-ClientTraceId' : get_guid (),
        }
    };

    let req = https.request (request_params, response_handler);
    req.write (content);
    req.end ();
}
let content = JSON.stringify ([{'Text' : text}]);
Translate (content);
}

var other_lang=false;
var lang_other=""
bot.dialog('speak in other languages', [
     function (session) {
        other_lang=true;
        builder.Prompts.text(session, 'Which language you want to speak in?');
    },
    function (session, results,next) {
        lang_other =results.response;
        next();
    },
    function (session,next) {
        session.sendTyping();
        setTimeout(() => {
            if(Translate_lang_flag){
                if(languageTo!="")
                translateMe(session,message,languageTo);
                else
                langtranslate(session)
                setTimeout(() => {
           Translate_lang_flag=false;
        }, 10);}
            else
            langtranslate(session)
        }, 1000);
    }
]).triggerAction({ matches: /speak in other languages/i });


//=====================================
var que=""
var other_lang=false;
var lang_other="";

bot.dialog('set language', [
     function (session) {
        other_lang=true;
        builder.Prompts.text(session, 'Which language you want to speak in?');
    },
    function (session, results) {
        lang_other =results.response;
        session.beginDialog('Start Again');
    }
]).triggerAction({ matches: /set language/i });

bot.dialog('set default language', [
     function (session) {
        other_lang=false;
        session.send("OK English is default");
        session.beginDialog('Introduce');
    }
]).triggerAction({ matches: /set default language/i });






function translateQue(session,textToTranslate) {

let fs = require ('fs');
let https = require ('https');
let subscriptionKey = 'b7bfbd192aa5442eb355e828444ace35';
let host = 'api.cognitive.microsofttranslator.com';
let path = '/translate?api-version=3.0';
let params = '&to=en';
let text = textToTranslate;
let response_handler = function (response) {
    let body = '';
    response.on ('data', function (d) {
        body += d;
    });
    response.on ('end', function () {
        let json = JSON.stringify(JSON.parse(body), null, 4);
        que= JSON.parse(body)[0].translations[0].text;
    });
    response.on ('error', function (e) {
        console.log ('Error: ' + e.message);
    });
};

let get_guid = function () {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

let Translate = function (content) {
    let request_params = {
        method : 'POST',
        hostname : host,
        path : path + params,
        headers : {
            'Content-Type' : 'application/json',
            'Ocp-Apim-Subscription-Key' : subscriptionKey,
            'X-ClientTraceId' : get_guid (),
        }
    };

    let req = https.request (request_params, response_handler);
    req.write (content);
    req.end ();
}
let content = JSON.stringify ([{'Text' : text}]);
Translate (content);
}


function translateAns(session,textToTranslate,p) {
let fs = require ('fs');
let https = require ('https');
let subscriptionKey = 'b7bfbd192aa5442eb355e828444ace35';
let host = 'api.cognitive.microsofttranslator.com';
let path = '/translate?api-version=3.0';
let params = '&to='+p;
let text = textToTranslate;
let response_handler = function (response) {
    let body = '';
    response.on ('data', function (d) {
        body += d;
    });
    response.on ('end', function () {
        let json = JSON.stringify(JSON.parse(body), null, 4);
        session.send(JSON.parse(body)[0].translations[0].text);
    });
    response.on ('error', function (e) {
        console.log ('Error: ' + e.message);
    });
};

let get_guid = function () {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

let Translate = function (content) {
    let request_params = {
        method : 'POST',
        hostname : host,
        path : path + params,
        headers : {
            'Content-Type' : 'application/json',
            'Ocp-Apim-Subscription-Key' : subscriptionKey,
            'X-ClientTraceId' : get_guid (),
        }
    };

    let req = https.request (request_params, response_handler);
    req.write (content);
    req.end ();
}
let content = JSON.stringify ([{'Text' : text}]);
Translate (content);
}



function translateSug(session,textToTranslate,p) {

let fs = require ('fs');
let https = require ('https');
let subscriptionKey = 'b7bfbd192aa5442eb355e828444ace35';
let host = 'api.cognitive.microsofttranslator.com';
let path = '/translate?api-version=3.0';
let params = '&to='+p;
let text = textToTranslate;
let response_handler = function (response) {
    let body = '';
    response.on ('data', function (d) {
        body += d;
    });
    response.on ('end', function () {
        let json = JSON.stringify(JSON.parse(body), null, 4);
        que= JSON.parse(body)[0].translations[0].text;
    });
    response.on ('error', function (e) {
        console.log ('Error: ' + e.message);
    });
};

let get_guid = function () {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

let Translate = function (content) {
    let request_params = {
        method : 'POST',
        hostname : host,
        path : path + params,
        headers : {
            'Content-Type' : 'application/json',
            'Ocp-Apim-Subscription-Key' : subscriptionKey,
            'X-ClientTraceId' : get_guid (),
        }
    };

    let req = https.request (request_params, response_handler);
    req.write (content);
    req.end ();
}
let content = JSON.stringify ([{'Text' : text}]);
Translate (content);
}




















































var cookies = []

 

var emails = []





//======================================================================================================================

//---------------------------------------BOT DIALOGUES--------------------------------------------------------------

 

//FIRST---------------

bot.dialog('get my mails', [

 

    (session, args, next) => {

 

        if (!cookies[3]) {

 

               session.beginDialog('signinPrompt');

 

        } else {

 

            next();

 

        }

 

    },

 

    (session, results, next) => {

 

        if (cookies[3]) {

            var input = ["email", "calendar", "contacts", "quit", "logout"], options = ['Get Mails', 'Get Events', 'Get Contacts', 'Quit', 'LogOut'];

            // They're logged in

            session.send('Welcome ' + cookies[3] + "." + " How can I help you?");

 

            // builder.Prompts.text(session,  "* To get the latest Emails, type 'email'.\n\n* To get Calendar Events type 'calendar',\n\n* For Contacts type 'contacts'\n\n* To Quit, type 'quit'. \n\n* To Log Out, type 'logout'. ");

 

            // create the card based on selection

 

            sendoptionCard(session, input, options);

            // builder.Prompts.text(session, msg) ;




        } else {

 

            session.endConversation("Goodbye.");

 

        }

 

    },

 

    (session, results, next) => {

        console.log('results...' + results);

        var resp = results.response;

 

        if (resp === 'Show my mails') {

 

            // session.beginDialog('workPrompt');

 

            session.beginDialog('sendMails');

 

        } else if (resp === 'Show me the calendar events') {

 

            session.beginDialog('calendar');

 

        } else if (resp === 'Show my contacts') {

 

            session.beginDialog('contacts');

 

        } else if (resp === 'quit') {

 

            session.endConversation("Goodbye.");

 

        } else if (resp === 'logout') {

            cookies = [];

            session.userData.loginData = null;

 

            session.userData.userName = null;

 

            session.userData.accessToken = null;

 

            session.userData.refreshToken = null;

 

            session.endConversation("You have logged out. Goodbye.");

 

        } else {

 

            next();

 

        }

 

    },

 

    (session, results) => {

 

        session.replaceDialog('/');

 

    }

 

]);

 

//SECOND=========================================

 

bot.dialog('signinPrompt', [

 

    (session, args) => {

        login(session);

    },

 

    (session, results) => {

 

        if (results.response === `login`) {

 

            // session.beginDialog('validateCode');

            if (cookies[0]) {

 

                session.endDialogWithResult({ response: true });

 

            } else {

 

                session.send("hmm... Looks like that was an invalid code. Please try again.");

 

                session.replaceDialog('signinPrompt', { invalid: true });

 

            }

 

        } else {

            session.send('Please type "login" again.')

            session.replaceDialog('signinPrompt', { invalid: true });

 

        }

 

    },

 

    (session, results) => {

 

        if (results.response) {

 

            session.endDialogWithResult({ response: true });

 

        } else {

 

            session.endDialogWithResult({ response: false });

 

        }

 

    }

 

]);

 

//===============================================   

 

// bot.dialog('validateCode', [

 

//     (session) => {

 

//         builder.Prompts.text(session, "Please type 'ok' to access outlook. ");

 

//     },

 

//     (session, results) => {

 

//         // const code = results.response;

//         const code = cookies[0]

 

//         console.log(code)

 

//         if (code == 'quit') {

 

//             session.endDialogWithResult({ response: false });

 

//         } else {

 

//             if (code == cookies[0]) {

 

//                 session.endDialogWithResult({ response: true });

 

//             } else {

 

//                 session.send("hmm... Looks like You are logged out. Please try again.");

 

//                 session.replaceDialog('validateCode');

 

//             }

 

//         }

 

//     }

 

// ]);

 

bot.dialog('sendMails', [

 

    (session, args, next) => {

 

        mail(session);

        session.send("Okay these are your latest recieved mails.");

 

    }

    // ,

 

    // (session, results) => {

 

    //     session.replaceDialog('/');

 

    // }

 

]);

 

bot.dialog('calendar', [

 

    (session, args) => {

 

        calendar(session);

        session.send(" Here's your outlook calendar events.");

 

    }

    // ,

 

    // (session, results) => {

 

    //     session.replaceDialog('/');

 

    // }

 

]);

 

bot.dialog('contacts', [

 

    (session, args) => {

 

        contacts(session);

       

    }

    // ,

 

    // (session, results) => {

 

    //     session.replaceDialog('/');

 

    // }

 

]);






//==============================FUNCTIONS===========================================================================================

 

//-------------------------------------------------------------------------------------------------------------------

//when signin button clicked in the bot ==> localhost 3000==>homepage

server.get("/", function home(response, request, next) {

    console.log('Request handler \'home\' was called.');

    response.writeHead(200, { 'Content-Type': 'text/html' });

    response.end();

    next();

});

 

//THIRD =====

 

function login(session) {

    var link = authHelper.getAuthUrl()

    var msg = new builder.Message(session)

        .attachments([

            new builder.SigninCard(session)

                .text("Welcome! Please click on the below link to access outlook.")

                .button("signin", link)

        ]);

    session.send(msg);

    builder.Prompts.text(session, "Please type 'login' to continue.");

}




bot.dialog('signin', [

 

    (session, results) => {

 

        console.log('signin callback: ' + results);

 

        session.endDialog();

 

    }

 

]);





server.get("/authorize", function authorize(response, request, next) {

 

    console.log('Request handler \'authorize\' was called.');

 

// console.log(response._url.query);
console.log("response "+response);
 

    // The authorization code is passed as a query parameter

    var url_parts = response._url.query;

 

    var code= url_parts.replace("code=","")
    var code_arr=code.split("&session_state=")
    code_arr[0]


    //console.log(url_parts)

// console.log("url part:"+ url_parts);

    console.log("Code "+code_arr[0])

 

    // console.log('Code: ' + code);

 

    authHelper.getTokenFromCode(code_arr[0], tokenReceived, response);

 

});




function tokenReceived(response, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
    } else {

        getUserEmail(token.token.access_token, function (error, email) {

 

            if (error) {

 

                console.log('getUserEmail returned an error: ' + error);

 

            } else if (email) {

 

                cookies = [token.token.access_token, token.token.refresh_token, token.token.expires_at.getTime(), email];

 

                // response.writeHead(302, { 'Location': 'http://localhost:8080/code' });

 

                // response.end();

 

            }

 

        });

 

    }

 

}




function getUserEmail(token, callback) {

 

    // Set the API endpoint to use the v2.0 endpoint

 

    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');




    // Set up oData parameters

 

    var queryParams = {

 

        '$select': 'DisplayName, EmailAddress',

 

    };




    outlook.base.getUser({ token: token, odataParams: queryParams }, function (error, user) {

 

        if (error) {

 

            callback(error, null);

 

        } else {

 

            callback(null, user.EmailAddress);

 

        }

 

    });

 

}

 

function getValueFromCookie(valueName, cookie) {

 

    if (cookie.indexOf(valueName) !== -1) {

 

        var start = cookie.indexOf(valueName) + valueName.length + 1;

 

        var end = cookie.indexOf(';', start);

 

        end = end === -1 ? cookie.length : end;

 

        return cookie.substring(start, end);

 

    }

 

}




function getAccessToken(request, response, callback) {

 

    var expiration = new Date(parseFloat(cookies[2]));




    if (expiration <= new Date()) {

 

        // refresh token

 

        console.log('TOKEN EXPIRED, REFRESHING');

 

        var refresh_token = cookies[1];

 

        authHelper.refreshAccessToken(refresh_token, function (error, newToken) {

 

            if (error) {

 

                callback(error, null);

 

            } else if (newToken) {

 

                cookies = [newToken.token.access_token, newToken.token.refresh_token, newToken.token.expires_at.getTime()];

 

                callback(null, newToken.token.access_token);

 

            }

 

        });

 

    } else {

 

        // Return cached token

 

        var access_token = cookies[0];

 

        callback(null, access_token);

 

    }

 

}




server.get("/code", function code(response, request) {

 

    getAccessToken(request, response, function (error, token) {

 

        console.log('Token found in cookie: ', token);

 

        var email = cookies[3]

 

        console.log('Email found in cookie: ', email);

 

        if (token) {

 

            response.writeHead(200, { 'Content-Type': 'text/html' });

 

            response.write('<div align="center"><h1>Welcome  ' + email + '</h1></div>');

 

            response.write("<div align='center'><h3>Please go back to the bot. You'll be able to access your Outlook Account now. </h3></div>");

            response.end();

 

        } else {

 

            response.writeHead(200, { 'Content-Type': 'text/html' });

 

            response.write('<p> No token found in cookie!</p>');

 

            response.end();

 

        }

 

    });

 

});




 function mail(session, response, request) {

 

    getAccessToken(request, response, function (error, token) {

 

        console.log('Token found in cookie: ', token);

 

        var email = cookies[3]

 

        console.log('Email found in cookie: ', email);

 

        if (token) {

 

            var queryParams = {

 

                '$select': 'Subject,ReceivedDateTime,From,IsRead, BodyPreview',

 

                '$orderby': 'ReceivedDateTime desc',

 

                '$top': 10

 

            };

 

            // Set the API endpoint to use the v2.0 endpoint

 

            outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

 

            // Set the anchor mailbox to the user's SMTP address

 

            outlook.base.setAnchorMailbox(email);

 

            outlook.mail.getMessages({ token: token, folderId: 'inbox', odataParams: queryParams },

 

                function (error, result) {

 

                    if (error) {

 

                        console.log('getMessages returned an error: ' + error);

 

                    }

 

                    else if (result) {

 

                        // console.log('getMessages returned ' + result.value.length + ' messages.');

 

                        // var i = 0;

                        // session.send("Okay Iam in your Inbox now.");

                        // result.value.forEach(function (message) {

 

                        //     console.log(' Subject: ' + message.Subject);

 

                        //     var from = message.From ? message.From.EmailAddress.Name : 'NONE';

 

                        //     emails[i] = "From :" + from + " Subject :" + message.Subject + " on " + message.ReceivedDateTime.toString();

 

                        //     session.send(emails[i]);

 

                        //     i++;

                        console.log('getMessages returned ' + result.value.length + ' messages.');

                        var i = 0, sub = [], tim = [], fromadd = [], body = [];

                        result.value.forEach(function (message) {

                            console.log(' Subject: ' + message.Subject);

                            console.log('message body:' + message.BodyPreview);

                            var from = message.From ? message.From.EmailAddress.Name : 'NONE';

                            sub[i] = message.Subject

                            tim[i] = message.ReceivedDateTime.toString();

                            fromadd[i] = from;

                            body[i] = message.BodyPreview;

                            // emails[i]="From :"+from+" Subject :"+message.Subject+" on "+message.ReceivedDateTime.toString();

                            // session.send(emails[i]);

                            i++;

                        });

                        sendCardMail(session, fromadd, tim, sub, body);

                        session.endDialogWithResult({

                            resumed: builder.ResumeReason.notCompleted

                            });

                           

                        // session.replaceDialog('/');

                    }

                });

 

        } else {

 

            console.log('No token found in cookie!');

 

        }

 

    });

 

}




function calendar(session, response, request) {

 

    var token = cookies[0];

 

    console.log('Token found in cookie: ', token);

 

    var email = cookies[3];

 

    console.log('Email found in cookie: ', email);

 

    if (token) {

 

        var queryParams = {

 

            '$select': 'Subject,Start,End,Attendees, BodyPreview',

 

            '$orderby': 'Start/DateTime desc',

 

            '$top': 10

 

        };

 

        // Set the API endpoint to use the v2.0 endpoint

 

        outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

 

        // Set the anchor mailbox to the user's SMTP address

 

        outlook.base.setAnchorMailbox(email);

 

        // Set the preferred time zone.

 

        // The API will return event date/times in this time zone.

 

        outlook.base.setPreferredTimeZone('Eastern Standard Time');

 

        outlook.calendar.getEvents({ token: token, odataParams: queryParams },

 

            function (error, result) {

 

                if (error) {

 

                    console.log('getEvents returned an error: ' + error);

 

                } else if (result) {

                    console.log('getEvents returned ' + result.value.length + ' events.');

 

                    var i = 0, sub = [], tim = [], attend = [], body = [];

                    result.value.forEach(function (event) {

                        console.log(' Subject: ' + event.Subject);

                        console.log(' Starting Time: ' + event.Start.DateTime.toString());

                        console.log(' Ending Time: ' + event.End.DateTime.toString());

                        console.log(' Attendees: ' + buildAttendeeString(event.Attendees));

                        console.log(' Event dump: ' + JSON.stringify(event));

                        body[i] = event.BodyPreview

                        sub[i] = event.Subject

                        tim[i] = event.Start.DateTime.toString() + ' to ' + event.End.DateTime.toString()

                        attend[i] = buildAttendeeString(event.Attendees);

                        i++;

                    });

                    sendCardCalendar(session, sub, tim, attend, body);

                    // session.replaceDialog('/');

                    session.endDialogWithResult({

                        resumed: builder.ResumeReason.notCompleted

                        });

                       

                }

            });

    }

    else {

 

        console.log('No token found in cookie!');

 

    }

 

}




function contacts(session, request, response) {

 

    var token = cookies[0]

 

    console.log('Token found in cookie: ', token);

 

    var email = cookies[3]

 

    console.log('Email found in cookie: ', email);

 

    if (token) {

 

        var queryParams = {

 

            '$select': 'GivenName,Surname,EmailAddresses',

 

            '$orderby': 'GivenName asc',

 

            '$top': 10

 

        };

 

        // Set the API endpoint to use the v2.0 endpoint

 

        outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

 

        // Set the anchor mailbox to the user's SMTP address

 

        outlook.base.setAnchorMailbox(email);

 

        outlook.contacts.getContacts({ token: token, odataParams: queryParams },

 

            function (error, result) {

 

                if (error) {

 

                    console.log('getContacts returned an error: ' + error);

                    session.replaceDialog('/');

 

                } else if (result) {

 

                    console.log('getContacts returned ' + result.value.length + ' contacts.');

                    var i = 0, firstName = [], lastName = [], mail = [];

                    result.value.forEach(function (contact) {

 

                        var email = contact.EmailAddresses[0] ? contact.EmailAddresses[0].Address : 'NONE';

 

                        console.log('First name: ' + contact.GivenName);

 

                        console.log('Last name: ' + contact.Surname);

 

                        console.log('Email: ' + email);

                        firstName[i] = contact.GivenName;

                        lastName[i] = contact.Surname;

                        mail[i] = email;

                        i++;

                    });

                    session.send("You have "+ result.value.length+" outlook contacts. Here they are...");

                    sendcardContacts(session, firstName, lastName, mail);

                    // session.send('First name: ' + contact.GivenName + ' Last name: ' + contact.Surname + ' Email: ' + email);

                

                    session.endDialogWithResult({

                        resumed: builder.ResumeReason.notCompleted

                        });

                      

 

                }

            });

 

    }

 

    else {

 

        console.log('No token found in cookie!');

        session.replaceDialog('/');

    }

 

}





function buildAttendeeString(attendees) {

 

    var attendeeString = 'wut';

 

    if (attendees) {

 

        attendeeString = '';




        attendees.forEach(function (attendee) {

 

            attendeeString += attendee.EmailAddress.Name + "<br>";




            // attendeeString += ' Type:' + attendee.Type;

 

            // attendeeString += ' Response:' + attendee.Status.Response;

 

            // attendeeString += ' Respond time:' + attendee.Status.Time;

 

        });

 

    }




    return attendeeString;

 

}

 

///===========================card attachment==================================

function sendCardMail(session, fromadd, tim, sub, body) {

    var attachments = [];

 

    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (sub[i] != null) {

 

        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "From: " + fromadd[i],

                        "size": "medium",

                        "weight": "bolder"

                    },

                    {

                        "type": "TextBlock",

                        "text": "Recieved at: " + tim[i],

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "Subject: ",

                        "size": "medium",

                        "weight": "bolder",

 

                    },

                    {

                        "type": "TextBlock",

                        "text": sub[i],

                        "size": "medium",

 

                        "wrap": true

                    },

 

                ],

                "actions": [

                    {

                        "type": "Action.ShowCard",

                        "title": "More...",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Content: ",

                                    "size": "medium",

                                    "wrap": true,

                                    "weight": "bolder"

                                },

                                {

                                    "type": "TextBlock",

                                    "width": "stretch",

                                    "height": "stretch",

                                    "text": body[i],

                                    "size": "medium",

                                    "wrap": true

                                }

                            ]

                        }

                    }

                ]

            }

        }

 

        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);

 

}

 

function sendCardCalendar(session, sub, tim, attend, body) {

    var attachments = [];

    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (sub[i] != null) {

        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "Subject: " + sub[i],

                        "size": "medium",

                        "weight": "bolder",

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "From: " + tim[i],

                        "wrap": true

                    }




                ],

                "actions": [

                    {

                        "type": "Action.ShowCard",

                        "title": "Event Details",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Event: ",

                                    "size": "medium",

                                    "weight": "bolder",

                                    "wrap": true,

 

                                },

                                {

                                    "type": "TextBlock",

                                    "text": body[i],

                                    "size": "medium",

                                    "wrap": true,

 

                                }

                            ]

                        }

                    },

                    {

                        "type": "Action.ShowCard",

                        "title": "View Attendies",

                        "card": {

                            "type": "AdaptiveCard",

                            "body": [

                                {

                                    "type": "TextBlock",

                                    "text": "Attendies",

                                    "size": "medium",

                                    "weight": "bolder",

                                    "wrap": true

                                },

                                {

                                    "type": "TextBlock",

                                    "text": attend[i],

                                    "size": "medium",

 

                                    "wrap": true

                                }

                            ]

                        }

                    }

                ]

            }

        }

        // var card = new builder.HeroCard(session)

        //     .title(sub[i])

        //     .subtitle(tim[i])

        //     .text(attend[i])

        //body[i]

        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);

}

 

function sendcardContacts(session, firstName, lastName, mail) {

    var attachments = [];

    var msg = new builder.Message(session);

    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    var i = 0

    while (firstName[i] != null) {

        var card = {

            'contentType': 'application/vnd.microsoft.card.adaptive',

            'content': {

                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',

                'type': 'AdaptiveCard',

                'version': '1.0',

                'body': [

                    {

                        "type": "TextBlock",

                        "text": "Name: " + firstName[i] + " " + lastName[i],

                        "size": "medium",

                        "weight": "bolder",

                        "wrap": true

                    },

                    {

                        "type": "TextBlock",

                        "text": "Email ID: " + mail[i],

                        "weight": "bolder",

                        "wrap": true

                    }

 

                ],

            }

        }

        // var card = new builder.HeroCard(session)

        //     .title(sub[i])

        //     .subtitle(tim[i])

        //     .text(attend[i])

        //body[i]

        attachments.push(card);

        i++;

    }

    msg.attachments(attachments)

    session.send(msg);

}

 

function sendoptionCard(session, input, options) {

    // console.log('im in')

    // var attachments = [];

    // var msg = new builder.Message(session);

    // msg.attachmentLayout(builder.AttachmentLayout.carousel);

 

    var i = 0

    while (input[i] != null) {

 

        // var card = new builder.HeroCard(session)

        //     .buttons([

        //         builder.CardAction.postBack(session, input[i], options[i])

 

        //     ])

        var msg = new builder.Message(session)

 
            .text("Choose any of the following.")
            .suggestedActions(

            builder.SuggestedActions.create(

                session, [

                    builder.CardAction.imBack(session, "Show my mails", "Get my mails"),

                    builder.CardAction.imBack(session, "Show me the calendar events", "View Events"),

                    builder.CardAction.imBack(session, "Show my contacts", "View Contacts"),

                    builder.CardAction.imBack(session, "logout", "Logout")

                ]

            ));

        //   session.send(msg);

 

        // attachments.push(card);

        i++;

    }

    // msg.attachments(attachments)

 

    builder.Prompts.text(session, msg);

 

}
