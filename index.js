var express = require('express'),
  builder = require('botbuilder'),
  bodyParser = require('body-parser'),
  request = require('request-promise'),
  fs = require('fs'),
  logger = require('./logger.js')('openchannel-skype4business'),
  moment = require('moment'),
  url = require('url'),
  path = require('path'),
  morgan = require('morgan'),
  Promise = require('bluebird'),
  mime = require('mime-types');

var app = express();

// Configuration
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
if (MOTION_URL) {
  var myUrl = url.parse(MOTION_URL);
  var DOMAIN = myUrl.protocol + '//' + myUrl.host;
}

// Microsoft Application Id
var APPLICATION_ID = config.microsoft_app_id;
// Microsoft Application Password
var APPLICATION_PASSWORD = config.microsoft_app_password;
// Credentials of the SkypeForBusiness user
var USERNAME = config.auth.username;
var PASSWORD = config.auth.password;

if (!(APPLICATION_ID && APPLICATION_PASSWORD && MOTION_URL && DOMAIN && USERNAME && PASSWORD)) {
  logger.error("Missing configuration values");
  process.exit(1);
}

//Start server
app.listen(PORT, function() {
  logger.info('Service listening on port ' + PORT);
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
  appId: APPLICATION_ID,
  appPassword: APPLICATION_PASSWORD
});

// Listen for messages from users 
app.post('/api/messages', connector.listen());

var address = null;
var inMemoryStorage = new builder.MemoryBotStorage();

// Initialization of UniversalBot 
var bot = new builder.UniversalBot(connector, function(session) {
  try {
    if (session && session.message && session.message.type === "message") {
      var msg = session.message;
      address = msg.address;
      session.userData.savedAddress = address;

      // Plain text message
      if (msg.textFormat === "plain") {
        return sendDataToMotion(address.user.id, address.user.name, msg.text, null);
      }

      // Message with attachment
      if (msg.attachments && msg.attachments.length > 0) {
        var attachment = msg.attachments[0];
        var myUrl = url.parse(attachment.contentUrl);
        var tempName = moment().unix() + path.extname(myUrl.pathname);
        var originalFilename = attachment.name;
        var w = fs.createWriteStream(__dirname + '/' + tempName);

        // Download attachment
        // Skype & MS Teams attachment URLs are secured by a JwtToken, so we need to pass the token from our bot.
        var fileDownload = checkRequiresToken(msg)
            ? requestWithToken(attachment.contentUrl)
            : request(attachment.contentUrl);
            
        fileDownload
        .pipe(w)
        .on('error', function(err){
          var errorMessage = 'Error getting attachment from SkypeForBusiness';
          logger.error(errorMessage);
          logger.error(err);
          // Send a message to Motion to warn that an error occured in sending the message
          return sendDataToMotion(address.user.id, address.user.name, errorMessage, null);

        })
        
        w.on('finish', function(){
          var uploadOptions = {
            method: 'POST',
            uri: DOMAIN + '/api/attachments',
            auth: {
              user: USERNAME,
              pass: PASSWORD
            },
            formData: {
              file: {
                value: fs.createReadStream(__dirname + '/' + tempName),
                options: {
                  filename: originalFilename
                }
              }
            },
            json: true
          };

        return request(uploadOptions)
          .then(function(attachment) {
            if(!attachment) {
              throw new Error('Unable to get uploaded attachment id!');
            }
            return sendDataToMotion(address.user.id, address.user.name, originalFilename, attachment.id);
          })
          .catch(function(err) {
            var errorMessage = 'Error uploading attachment to xCALLY Motion server';
            logger.error(errorMessage, err);
            deleteTempFile(__dirname + '/' + tempName);
            return sendDataToMotion(address.user.id, address.user.name, errorMessage, null);
          })
        });
      }
    }
  } catch (e) {
    logger.error(JSON.stringify(e));
  }
}).set("storage", inMemoryStorage);

//Send the message to Motion 
function sendDataToMotion(senderId, senderName, content, attachmentId){
    return request({
      method: 'POST',
      uri: config.url,
      body: {
        from: senderId,
        firstName: senderName,
        body: content,
        skype: senderId,
        AttachmentId: attachmentId,
        mapKey: "skype"
      },
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
  logger.info('Sending a message to SkypeForBusiness', req.body);
  if (address) {
    if(!req.body.body){
      return res.status(500).send({
        message: 'The body of the message cannot be null'
      });
    }
    return sendDataToSkype(req.body, res);
  };
});

function sendDataToSkype(data, res) {
  var message = new builder.Message().address(address);

  if(data.AttachmentId){
    if(!data.body) {
      var errorMessage = 'Unable to retrieve the attachment\'s name';
      logger.error(errorMessage)
      return res.status(500).send({
        message: errorMessage
      })
    }
    
    var fileExtension = path.extname(data.body);
    var filename = moment().unix() + fileExtension;
    var w = fs.createWriteStream(__dirname + '/' + filename);

    request({
      uri: DOMAIN + '/api/attachments/' + data.AttachmentId + '/download',
      method: 'GET',
      auth: {
        user: USERNAME,
        pass: PASSWORD
      }
    })
    .on('error', function(err){
      logger.error('An error occurred while retrieving the attachment from the Motion server' +
                   'Cannot send the message to %s', to)
      logger.error(err);
      return res.status(500).send(err);
    })
    .pipe(w);

    w.on('finish', function(){
      var attachment = {
        contentUrl: __dirname + '/' + filename,
        contentType: mime.lookup(fileExtension)
      }
      message.addAttachment(attachment);
      message.text(data.body);

      bot.send(message, function(err){
        if(err){
          logger.error(err);
          res.status(500).send(err);
        }
        else{
          // if (filename) {
          //   deleteTempFile(__dirname + '/' + filename);
          // }
          res.status(200).send('ok');
        }
      });
    })
  }
  else{
    message.text(data.body);
    bot.send(message, function(err){
      if(err){
        logger.error(err);
        res.status(500).send(err);
      }
      else{
        res.status(200).send('ok');
      }
    });
  }
};

// Helper methods
// Request file with Authentication Header
var requestWithToken = function (url) {
  return obtainToken().then(function (token) {
      return request({
          url: url,
          headers: {
              'Authorization': 'Bearer ' + token,
              'Content-Type': 'application/octet-stream'
          }
      });
  });
};

// Promise for obtaining JWT Token (requested once)
var obtainToken = Promise.promisify(connector.getAccessToken.bind(connector));

var checkRequiresToken = function (message) {
  return message.source === 'skype' || message.source === 'msteams';
};

//Delete the temporary file stored locally
function deleteTempFile(path) {
  fs.unlink(path, function(err) {
    if (err) {
      logger.error('Unable to delete the temporary file', err);
    } else {
      logger.debug('Temporary file correctly deleted!');
    }
  });
};
