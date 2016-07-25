var express = require('express'),
    SkypeAPI = require('skypeapi'),
    rp = require('request-promise'),
    bodyParser = require('body-parser'),
    fs = require('fs');

var app = express();

// Configuration
try {
    var configJSON = fs.readFileSync(__dirname + '/config.json');
    var config = JSON.parse(configJSON.toString());
} catch (e) {
    console.error('File config.json not found or is invalid: ' + e.message);
    process.exit(1);
}

var skype = new SkypeAPI({
    username: config.username,
    password: config.password
});

var getFrom = function(value) {
    return value.split('/').pop().split(':').pop();
};

var logger = function(value, type) {
    console.log('[' + new Date() + ']', '[OPENCHANNEL]', (type) ? '[' + type + ']' : '[INFO]', value);
};

skype.on('Chat', function(e) {
    if (getFrom(e.from) !== config.username) {
        var options = {
            method: 'POST',
            uri: config.url,
            body: {
                from: e.channel,
                body: e.content,
                name: e.imdisplayname
            },
            json: true
        };

        rp(options)
            .then(function(res) {
                // Process response...
                logger(res);
            })
            .catch(function(err) {
                // Crawling failed...
                logger(err, 'ERROR');
            });
    }
});

//bodyParser to get POST parameters.
app.use(bodyParser.urlencoded({
    extended: false
}));
app.use(bodyParser.json());

app.post('/sendMessage', function(req, res) {
    var recipientId = req.body.to;
    var message = req.body.body;
    // Send message to recipient
    skype.sendMessage(req.body.to, req.body.body);
    res.sendStatus(200);
});

var port = config.port || 3000;

//Start server
app.listen(port, function() {
    console.log('Service listening on port ' + port);
});
