var express = require('express'),
  builder = require('botbuilder'),
  bodyParser = require('body-parser'),
  request = require('request-promise'),
  fs = require('fs-extra'),
  logger = require('./logger.js')('openchannel-skype4business'),
  moment = require('moment'),
  morgan = require('morgan');

var app = express();

// Retrieve the configuration
try {
  var configJSON = fs.readFileSync(__dirname + '/config.json');
  var config = JSON.parse(configJSON.toString());
} catch (e) {
  logger.error('File config.json not found or invalid: ' + e.message);
  process.exit(1);
}

var PORT = config.port || 3003;
// URL that can be retrieved from the Motion OpenChannel Account
var MOTION_URL = config.url;

// Microsoft Application Id
var APPLICATION_ID = config.microsoft_app_id;
// Microsoft Application Password
var APPLICATION_PASSWORD = config.microsoft_app_password;

if (!(APPLICATION_ID && APPLICATION_PASSWORD && MOTION_URL)) {
  logger.error("Missing configuration values");
  process.exit(1);
}

// Start the server
app.listen(PORT, function () {
  logger.info('Service listening on port ' + PORT);
});

// bodyParser to get POST parameters
app.use(bodyParser.urlencoded({
  extended: false
}));
app.use(bodyParser.json());

morgan.token('remote-address', function (req, res) {
  return req.headers['x-forwarded-for'] ? req.headers['x-forwarded-for'] : req.connection.remoteAddress || req.ip;
});
morgan.token('datetime', function (req, res) {
  return moment().format('YYYY-MM-DD HH:mm:ss');
});

app.use(morgan('VERBOSE [:datetime] [REQUEST] [OPENCHANNEL-SKYPE4BUSINESS] - :method :remote-address :remote-user :url :status :response-time ms - :res[content-length]'));

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
  appId: APPLICATION_ID,
  appPassword: APPLICATION_PASSWORD
});

app.get('/status', function (req, res) {
  res.send('The service is running correctly');
});

// Listen for messages from users 
app.post('/api/messages', connector.listen());

var inMemoryStorage = new builder.MemoryBotStorage();
var address = null;

// Initialization of UniversalBot 
var bot = new builder.UniversalBot(connector, function (session) {
  try {
    if (session && session.message && session.message.type === "message") {
      var msg = session.message;
      address = msg.address;

      // Plain text message
      if (msg.textFormat === "plain") {
        // Check if the message was supposed to be an attachment
        if (!msg.text.startsWith("Application-Name: File Transfer")) {
          return sendDataToMotion(address.user.id, address.user.name, msg.text, null, address.conversation.id);
        }
      }
    }
  } catch (e) {
    logger.error(JSON.stringify(e));
  }
}).set("storage", inMemoryStorage);

// Send the message to Motion 
function sendDataToMotion(senderId, senderName, content, attachmentId, conversationId) {
  return request({
      method: 'POST',
      uri: config.url,
      body: {
        from: senderId,
        firstName: senderName,
        body: content,
        skype: senderId,
        AttachmentId: attachmentId,
        mapKey: "skype",
        threadId: conversationId
      },
      json: true
    })
    .then(function (res) {
      logger.info(res);
    })
    .catch(function (err) {
      logger.error(JSON.stringify(err));
    });
};

app.post('/sendMessage', function (req, res) {
  logger.info('Sending a message to SkypeForBusiness', req.body);
  if (address) {
    if (!req.body.body) {
      return res.status(500).send({
        message: 'The body of the message cannot be null'
      });
    }
    return sendDataToSkype(address, req.body, res);
  } else {
    errorMessage = 'Cannot send the message. The conversation might have been closed or deleted.';
    return res.status(500).send({
      message: errorMessage
    });
  }
});

// Send message to SkypeForBusiness
function sendDataToSkype(address, data, res) {
  var message = new builder.Message().address(address);
  if (data.AttachmentId) {
    var warningMessage = 'Attachments are not supported';
    return res.status(501).send({
      message: warningMessage
    });
  } else {
    message.text(data.body);
    bot.send(message, function (err) {
      if (err) {
        logger.error(err);
        res.status(500).send(err);
      } else {
        res.status(200).send('ok');
      }
    });
  }
};
