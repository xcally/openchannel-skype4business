var express = require('express'),
  builder = require('botbuilder'),
  botbuilder_azure = require("botbuilder-azure"),
  bodyParser = require('body-parser'),
  request = require('request-promise'),
  fs = require('fs'),
  logger = require('./logger.js')('openchannel-skype4business'),
  moment = require('moment'),
  morgan = require('morgan');

var app = express();

// Configuration
try {
  var configJSON = fs.readFileSync(__dirname + '/config.json');
  var config = JSON.parse(configJSON.toString());
} catch (e) {
  logger.error('File config.json not found or invalid: ' + e.message);
  process.exit(1);
}

var port = config.port || 3003;

//Start server
app.listen(port, function() {
  logger.info('Service listening on port ' + port);
});
  
//bodyParser to get POST parameters
app.use(bodyParser.urlencoded({
  extended: false
}));
app.use(bodyParser.json());

morgan.token('remote-address', function(req, res) {
  return req.headers['x-forwarded-for'] ? req.headers['x-forwarded-for'] : req.connection.remoteAddress || req.ip;
});
morgan.token('datetime', function(req, res) {
  return moment().format('YYYY-MM-DD HH:mm:ss');
});

app.use(morgan('VERBOSE [:datetime] [REQUEST] [OPENCHANNEL-SKYPE4BUSINESS] - :method :remote-address :remote-user :url :status :response-time ms - :res[content-length]'));

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
  appId: config.microsoft_app_id,
  appPassword: config.microsoft_app_password
});

// Listen for messages from users 
app.post('/api/messages', connector.listen());

var address = null;
var inMemoryStorage = new builder.MemoryBotStorage();

var bot = new builder.UniversalBot(connector, function (session) {
  try {
    if (session && session.message &&
      session.message.type === 'message' && session.message.textFormat === 'plain') {

      address = session.message.address;
      session.userData.savedAddress = address;

      if (address) {
        var data = {
          from: session.message.address.user.id,
          body: session.message.text,
          firstName: session.message.address.user.name,
          skype: session.message.address.user.id,
          mapKey: 'skype'
        };

        return sendData(data);
      }
    }
  } catch (e) {
    logger.error(JSON.stringify(e));
  }
}).set('storage', inMemoryStorage);

function sendData(data){
    return request({
      method: 'POST',
      uri: config.url,
      body: data,
      json: true
    })
    .then(function(res){
      logger.info(res);
    })
    .catch(function(err) {
      logger.error(JSON.stringify(err));
    });
};

app.post('/sendMessage', function(req, res) {
  try {
    logger.info('sendMessage', req.body);
      if (address) {
        sendProactiveMessage(req.body.body);
        res.status(200).send('ok');
      };
  } catch (e) {
    logger.error('Reply error:', JSON.stringify(e));
  }
});

function sendProactiveMessage(message) {
    var msg = new builder.Message().address(address);
    msg.text(message);

    bot.send(msg, function(err){
      if(err) logger.error(err);
    });
};
