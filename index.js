var express = require('express'),
    builder = require('botbuilder'),
    bodyParser = require('body-parser'),
    request = require('request-promise'),
    fs = require('fs-extra'),
    logger = require('./logger.js')('openchannel-skype4business'),
    moment = require('moment'),
    url = require('url'),
    path = require('path'),
    morgan = require('morgan'),
    Promise = require('bluebird'),
    mime = require('mime-types');

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
if (MOTION_URL) {
    var myUrl = url.parse(MOTION_URL);
    var DOMAIN = myUrl.protocol + '//' + myUrl.host;
}

var USERNAME = config.auth.username;
var PASSWORD = config.auth.password;

// Microsoft Application Id
var APPLICATION_ID = config.microsoft_app_id;
// Microsoft Application Password
var APPLICATION_PASSWORD = config.microsoft_app_password;

// Token for proxy
if (config.hasOwnProperty('proxy_token')) {
    var PROXY_TOKEN = config.proxy_token;
    if (!config.hasOwnProperty('proxy_url')) {
        logger.error('Missing proxy configuration values');
        process.exit(1);
    }
    var PROXY_URL = config.proxy_url;
}

if (!(APPLICATION_ID && APPLICATION_PASSWORD && MOTION_URL && USERNAME && PASSWORD)) {
    logger.error('Missing configuration values');
    process.exit(1);
}

var tempDirPath = '/tmp/attachments/';

// Create the temporary file that will store the conversations' address info
fs.ensureFile(__dirname + '/tmp/conversations_history.json')
    .then(function () {
        // Create the temporary directory that will store the attachments sent from Motion
        fs.ensureDir(__dirname + tempDirPath)
            .then(() => {
                logger.log('Temporary attachments folder created!')
            })
            .catch(err => {
                logger.error('An error occured while creating the temporary attachments folder', err)
            })

        fs.readJson(__dirname + '/tmp/conversations_history.json')
            .catch(function () {
                // Initialization of empty array
                var data = {
                    conversations: []
                }
                fs.writeJson(__dirname + '/tmp/conversations_history.json', data)
                    .then(function () {
                        logger.log('Temporary conversations history file created!')
                    })
                    .catch(function (err) {
                        logger.error('Cannot write to the temporary conversations history file', err)
                    })
            })
    })
    .catch(function (err) {
        logger.error('An error occured while creating the temporary conversations history file', err)
    })

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

// Return the service status
app.get('/api/status', function (req, res) {
    res.send('The service is running correctly');
});

// Listen for messages from users
app.post('/api/messages', connector.listen());

var inMemoryStorage = new builder.MemoryBotStorage();

// Initialization of UniversalBot
var bot = new builder.UniversalBot(connector, function (session) {
    try {
        if (session && session.message && session.message.type === 'message') {
            var msg = session.message;
            var address = msg.address;

            updateConversationsHistory(address);

            // Plain text message
            if (msg.textFormat === 'plain') {
                // Check if the message was supposed to be an attachment
                if (msg.text.startsWith("Application-Name: File Transfer")) {
                    logger.info('Cannot send attachment', msg.text);
                } else {
                    return sendDataToMotion(address.user.id, address.user.name, msg.text, null, address.conversation.id);
                }
            }
        }
    } catch (e) {
        logger.error(JSON.stringify(e));
    }
}).set('storage', inMemoryStorage);

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
                mapKey: 'skype',
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

// Listen for download requests
app.get('/api/attachments/:filename/download', function (req, res) {
    if (req.params.filename) {
        var downloadPath = path.join(__dirname, tempDirPath, req.params.filename);
        var extension = path.extname(req.params.filename);
        var type = mime.lookup(extension);

        logger.info(downloadPath);

        res.setHeader('Content-disposition', 'attachment; filename=' + req.params.id);
        res.setHeader('Content-type', type);
        res.download(downloadPath);
    }
});

app.post('/sendMessage', function (req, res) {
    logger.info('Sending a message to Skype For Business', req.body);
    var conversation = req.body.Interaction;
    var address = getConversationAddress(conversation.threadId);
    if (address) {
        if (!req.body.body) {
            return res.status(500).send({
                message: 'The body of the message cannot be null'
            });
        }
        return sendDataToSFB(address, req.body, res);
    } else {
        errorMessage = 'Cannot send the message. The conversation might have been closed or deleted.';
        return res.status(500).send({
            message: errorMessage
        });
    }
});

// Send message to Skype For Business
function sendDataToSFB(address, data, res) {
    var message = new builder.Message().address(address);
    if (data.AttachmentId) {
        if (!data.body) {
            var errorMessage = 'Unable to retrieve the attachment\'s name';
            logger.error(errorMessage)
            return res.status(500).send({
                message: errorMessage
            })
        }

        var fileExtension = path.extname(data.body);
        var filename = moment().unix() + fileExtension;
        var w = fs.createWriteStream(__dirname + tempDirPath + filename);

        // Retrieve the attachment from Motion
        request({
                uri: DOMAIN + '/api/attachments/' + data.AttachmentId + '/download',
                method: 'GET',
                auth: {
                    user: USERNAME,
                    pass: PASSWORD
                }
            })
            .on('error', function (err) {
                logger.error('An error occurred while retrieving the attachment from the Motion server' +
                    'Cannot send the message to %s', to)
                logger.error(err);
                return res.status(500).send(err);
            })
            .pipe(w);

        w.on('finish', function () {
            var attachment = {
                contentUrl: path.join(__dirname, tempDirPath, filename),
                contentType: mime.lookup(fileExtension)
            }

            // Create url to download the file
            var downloadPath = path.join((PROXY_URL) ? PROXY_URL : DOMAIN, 'skype4business/api/attachments', filename, 'download');
            if (PROXY_URL) {
                downloadPath = downloadPath + '?token=' + PROXY_TOKEN;
            }
            var msg = '<a href="' + downloadPath + '">' + data.body + '</a>';

            message.text(msg);
            bot.send(message, function (err) {
                if (err) {
                    logger.error(err);
                    res.status(500).send(err);
                } else {
                    res.status(200).send('ok');
                }
            });
        })
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

// Helper methods
// Update the temporary conversation history file
function updateConversationsHistory(address) {
    fs.readJson(__dirname + '/tmp/conversations_history.json')
        .then(function (historyJSON) {
            var exists = false;
            // Check if the current conversationId has already been stored
            for (var i = 0; i < historyJSON.conversations.length; i++) {
                if (historyJSON.conversations[i].conversation.id == address.conversation.id) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                // Add the conversation address
                historyJSON.conversations.push(address);
                fs.writeJson(__dirname + '/tmp/conversations_history.json', historyJSON)
                    .then(function () {
                        logger.log('Temporary conversations history file updated!')
                    })
                    .catch(function (err) {
                        logger.error(err)
                    })
            }
        })
        .catch(function (err) {
            logger.error(err)
        })
};

// Retrieve the conversation address
function getConversationAddress(addressId) {
    var historyJSON = fs.readJsonSync(__dirname + '/tmp/conversations_history.json')
    for (var i = 0; i < historyJSON.conversations.length; i++) {
        if (historyJSON.conversations[i].conversation.id == addressId) {
            return historyJSON.conversations[i];
        }
    }
};

// Delete the temporary file stored locally
function deleteTempFile(path) {
    fs.unlink(path, function (err) {
        if (err) {
            logger.error('Unable to delete the temporary file', err);
        } else {
            logger.debug('Temporary file correctly deleted!');
        }
    });
};
