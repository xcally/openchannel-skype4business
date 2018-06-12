var moment = require('moment');
var winston = require('winston');
var util = require('util');
var path = require('path');
var _ = require('lodash');

function filename(service, level) {
  return path.join('/var/log/xcally', util.format('%s-%s.log', service, level));
}

function formatter(service, options) {
  return util.format('VERBOSE [%s] [%s] [%s] - %s %s', moment().format("YYYY-MM-DD HH:mm:ss"), options.level.toUpperCase(), service.toUpperCase(), options.message, _.isEmpty(options.meta) ? '' : JSON.stringify(options.meta));
}

module.exports = function(service) {
  return new(winston.Logger)({
    transports: [
      new(winston.transports.File)({
        name: util.format('%s-%s', service, 'error'),
        filename: filename(service, 'error'),
        level: 'error',
        json: false,
        maxsize: 5242880,
        maxFiles: 7,
        tailable: true, //when the log file is full and the logging continues to another file, this new file will be named as the original one and the full one will have a consecutive number added to the name
        formatter: function(options) {
          return formatter(service, options);
        }
      }),
      new(winston.transports.File)({
        name: util.format('%s-%s', service, 'info'),
        filename: filename(service, 'info'),
        level: 'info',
        json: false,
        maxsize: 5242880,
        maxFiles: 7,
        tailable: true,
        formatter: function(options) {
          return formatter(service, options);
        }
      }),
      new(winston.transports.Console)({
        level: 'info',
        colorize: true,
        json: false,
        formatter: function(options) {
          return formatter(service, options);
        }
      })
    ]
  });
};
